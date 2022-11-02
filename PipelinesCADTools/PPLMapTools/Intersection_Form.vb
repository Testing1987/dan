Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Intersection_Form

    Dim Data_table_Layers As System.Data.DataTable
    Dim Freeze_operations As Boolean = False
    Dim Data_table_station_equation As System.Data.DataTable

    Private Sub ButtonTotal_length_of_reroutes_Click(sender As Object, e As EventArgs) Handles ButtonTotal_length_of_reroutes.Click

        If Freeze_operations = False Then
            Try


                Dim Data_Table_information As New System.Data.DataTable

                Data_Table_information.Columns.Add("LEN_GRID", GetType(Double))


                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions

                    Object_Prompt.MessageForAdding = vbLf & "Select one POLYLINE (2d OR 3d) (all objects IN THE SAME LAYER WILL BE ADDED TO THE TOTAL LENGTH):"


                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)





                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                Dim Col_ObjId_CL As New ObjectIdCollection


                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)

                                Dim Index_row As Double = 0



                                If TypeOf Ent1 Is Polyline3d Then

                                    Col_ObjId_CL.Add(Obj1.ObjectId)
                                End If


                                If TypeOf Ent1 Is Polyline Then
                                    Col_ObjId_CL.Add(Obj1.ObjectId)
                                End If

                                If Col_ObjId_CL.Count = 0 Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If





                                Dim Ent_CL As Entity = Trans1.GetObject(Col_ObjId_CL(0), OpenMode.ForRead)

                                Dim Total_Length As Double = 0

                                Dim ID_db As Double = 0

                                For Each ObjID In BTrecord
                                    Dim Ent_curve As Entity = Trans1.GetObject(ObjID, OpenMode.ForRead)
                                    If Ent_CL.Layer = Ent_curve.Layer Then
                                        If TypeOf (Ent_curve) Is Polyline Then
                                            Dim Curve1 As Polyline = Ent_curve
                                            Total_Length = Total_Length + Curve1.Length
                                            Data_Table_information.Rows.Add()
                                            Data_Table_information.Rows(ID_db).Item("LEN_GRID") = Curve1.Length
                                            ID_db = ID_db + 1
                                        End If
                                        If TypeOf (Ent_curve) Is Polyline3d Then
                                            Dim Curve1 As Polyline3d = Ent_curve
                                            Total_Length = Total_Length + Curve1.Length
                                            Data_Table_information.Rows.Add()
                                            Data_Table_information.Rows(ID_db).Item("LEN_GRID") = Curve1.Length
                                            ID_db = ID_db + 1
                                        End If
                                    End If
                                Next

                                MsgBox(Round(Total_Length, 2))

                                Editor1.Regen()
                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                If Data_Table_information.Rows.Count > 0 Then
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                    For i = 0 To Data_Table_information.Columns.Count - 1
                        W1.Cells(1, i + 1).value2 = Data_Table_information.Columns(i).ColumnName
                    Next
                    Dim Rand_excel As Double = 2
                    For i = 0 To Data_Table_information.Rows.Count - 1
                        For j = 0 To Data_Table_information.Columns.Count - 1
                            If IsDBNull(Data_Table_information.Rows(i).Item(j)) = False Then
                                W1.Cells(Rand_excel, j + 1).value2 = Data_Table_information.Rows(i).Item(j)
                            End If
                        Next
                        Rand_excel = Rand_excel + 1
                    Next
                End If

                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button_output_offset_Click(sender As Object, e As EventArgs) Handles Button_output_offset.Click
        If Freeze_operations = False Then

            Try



                If IsNumeric(TextBox_buffer.Text) = False Then
                    MsgBox("Please specify the Buffer size!")
                    Exit Sub
                End If


                Dim Col_station As String = "STATION"
                Dim Col_station_eq As String = "STATION_EQ"
                Dim Col_description As String = "DESCRIPTION"
                Dim Col_x As String = "X"
                Dim Col_y As String = "Y"
                Dim Col_z As String = "Z"
                Dim Col_z_point As String = "POINT_OBJECT_ELEVATION"
                Dim Col_layer_name As String = "LAYER_NAME"
                Dim Col_layer_description As String = "LAYER_DESCRIPTION"
                Dim Col_offset As String = "OFFSET"
                Dim Col_block_name As String = "BLOCK_NAME"
                Dim Col_left_right As String = "LEFT_RIGHT"
                Dim No_plot As String = "NO PLOT"




                Dim Data_Table_information As New System.Data.DataTable
                Data_Table_information.Columns.Add(Col_station, GetType(Double))
                Data_Table_information.Columns.Add(Col_station_eq, GetType(Double))
                Data_Table_information.Columns.Add(Col_description, GetType(String))

                Data_Table_information.Columns.Add(Col_x, GetType(Double))
                Data_Table_information.Columns.Add(Col_y, GetType(Double))
                Data_Table_information.Columns.Add(Col_z, GetType(Double))
                Data_Table_information.Columns.Add(Col_z_point, GetType(Double))


                Data_Table_information.Columns.Add(Col_layer_name, GetType(String))
                Data_Table_information.Columns.Add(Col_layer_description, GetType(String))


                Data_Table_information.Columns.Add(Col_offset, GetType(Double))
                Data_Table_information.Columns.Add(Col_block_name, GetType(String))
                Data_Table_information.Columns.Add(Col_left_right, GetType(String))



                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select centerline:"

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer(No_plot, 40, No_plot, False)

                                Dim Poly3d As Polyline3d = Nothing
                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)





                                Dim Index_row As Double = 0

                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection


                                Dim Poly2D As New Polyline


                                If TypeOf Ent1 Is Polyline3d Then
                                    Poly3d = Ent1
                                    Dim Index2d As Double = 0
                                    For Each ObjId As ObjectId In Poly3d
                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                        Index2d = Index2d + 1
                                    Next
                                End If


                                If TypeOf Ent1 Is Polyline Then
                                    Poly2D = Ent1.Clone
                                End If

                                If Poly2D = Nothing Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Poly2D.Elevation = 0




                                Dim Data_table_Columns_object_data As New System.Data.DataTable
                                Data_table_Columns_object_data.Columns.Add("NAME", GetType(String))
                                Data_table_Columns_object_data.Columns.Add("NR", GetType(Integer))
                                Dim nrdec As Integer = 3
                                If CheckBox_zero_decimals.Checked = True Then nrdec = 0
                                Dim Segment_No As Integer = 1


                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables

                                For Each ObjID In BTrecord

                                    Dim Ent_Object As Entity = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                                    If Not Ent_Object = Nothing Then
                                        If Not Ent_Object.ObjectId = Ent1.ObjectId Then
                                            If TypeOf Ent_Object Is BlockReference Then
                                                Dim Block1 As BlockReference = Ent_Object

                                                Dim Point_on_poly2d As New Point3d
                                                Dim Point_on_poly3d As New Point3d
                                                Dim Point_on_poly As New Point3d

                                                Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(Block1.Position.X, Block1.Position.Y, 0), Vector3d.ZAxis, True)

                                                Dim Offset1 As Double = New Point3d(Block1.Position.X, Block1.Position.Y, 0).GetVectorTo(Point_on_poly2d).Length

                                                If Offset1 <= CDbl(TextBox_buffer.Text) Then

                                                    Data_Table_information.Rows.Add()

                                                    Dim ChainageGrid0 As Double

                                                    If Not Poly3d = Nothing Then
                                                        Dim Param2d As Double = Poly2D.GetParameterAtPoint(Point_on_poly2d)
                                                        Point_on_poly3d = Poly3d.GetPointAtParameter(Param2d)
                                                        Point_on_poly = Point_on_poly3d
                                                        ChainageGrid0 = Poly3d.GetDistAtPoint(Point_on_poly3d)
                                                    Else
                                                        Point_on_poly = Point_on_poly2d
                                                        ChainageGrid0 = Poly2D.GetDistAtPoint(Point_on_poly2d)
                                                    End If


                                                    Data_Table_information.Rows(Index_row).Item(Col_x) = Round(Point_on_poly.X, 3)
                                                    Data_Table_information.Rows(Index_row).Item(Col_y) = Round(Point_on_poly.Y, 3)
                                                    Data_Table_information.Rows(Index_row).Item(Col_z) = Round(Point_on_poly.Z, 3)
                                                    Data_Table_information.Rows(Index_row).Item(Col_station) = Round(ChainageGrid0, nrdec)
                                                    Data_Table_information.Rows(Index_row).Item(Col_station_eq) = Round(ChainageGrid0 + Get_equation_value(ChainageGrid0), nrdec)

                                                    Dim Valoare As String = Block1.Layer
                                                    Dim Valoare1 As String = ""
                                                    If IsNothing(Data_table_Layers) = False Then
                                                        If Data_table_Layers.Rows.Count > 0 Then
                                                            For i = 0 To Data_table_Layers.Rows.Count - 1
                                                                If IsDBNull(Data_table_Layers.Rows(i).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(i).Item(Col_description)) = False Then
                                                                    If Data_table_Layers.Rows(i).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                        Valoare1 = Data_table_Layers.Rows(i).Item(Col_description)
                                                                        Exit For
                                                                    End If

                                                                End If
                                                            Next
                                                        End If
                                                    End If

                                                    Data_Table_information.Rows(Index_row).Item(Col_layer_name) = Block1.Layer
                                                    If Not Valoare1 = "" Then
                                                        Data_Table_information.Rows(Index_row).Item(Col_layer_description) = Valoare1
                                                    End If




                                                    Data_Table_information.Rows(Index_row).Item(Col_offset) = Round(Offset1, nrdec)
                                                    Data_Table_information.Rows(Index_row).Item(Col_block_name) = get_block_name(Block1)
                                                    Dim Left_right As String = Angle_left_right(Poly2D, New Point3d(Block1.Position.X, Block1.Position.Y, 0), 1)
                                                    Data_Table_information.Rows(Index_row).Item(Col_left_right) = Left_right


                                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Block1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                        If IsNothing(Records1) = False Then
                                                            If Records1.Count > 0 Then
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

                                                                        If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                            Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                        End If
                                                                        If Not Replace(Valoare_field, " ", "") = "" Then
                                                                            Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If
                                                        End If
                                                    End Using

                                                    If Block1.AttributeCollection.Count > 0 Then

                                                        For Each id As ObjectId In Block1.AttributeCollection
                                                            If Not id.IsErased Then
                                                                Dim attRef As AttributeReference = DirectCast(Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                                                Dim Continut As String = attRef.TextString
                                                                Dim Tag As String = attRef.Tag
                                                                If Data_Table_information.Columns.Contains(Tag) = False Then
                                                                    Data_Table_information.Columns.Add(Tag, GetType(String))
                                                                End If
                                                                If Not Replace(Continut, " ", "") = "" Then
                                                                    Data_Table_information.Rows(Index_row).Item(Tag) = Continut
                                                                End If

                                                            End If
                                                        Next


                                                    End If

                                                    Index_row = Index_row + 1




                                                    Dim new_leader As New MLeader


                                                    Dim Station As String
                                                    Dim Station_EQ As String
                                                    If CheckBox_US_station.Checked = False Then
                                                        Station = Get_chainage_from_double(ChainageGrid0, 3)
                                                        Station_EQ = Get_chainage_from_double(ChainageGrid0 + Get_equation_value(ChainageGrid0), 3)
                                                    Else
                                                        Station = Get_chainage_feet_from_double(ChainageGrid0, 0)
                                                        Station_EQ = Get_chainage_feet_from_double(ChainageGrid0 + Get_equation_value(ChainageGrid0), 0)
                                                    End If
                                                    Dim eXTRA1 As String = ""
                                                    If Not Get_equation_value(ChainageGrid0) = 0 Then
                                                        eXTRA1 = "St_eq = " & Station_EQ & vbCrLf
                                                    End If


                                                    new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly,
                                                                                                       "East = " & Round(Point_on_poly.X, 3) & vbCrLf &
                                                                                                       "North = " & Round(Point_on_poly.Y, 3) & vbCrLf &
                                                                                                       "Elev = " & Round(Point_on_poly.Z, 3) & vbCrLf &
                                                                                                       "Station = " & Station & vbCrLf & eXTRA1 &
                                                                                                       "Offset = " & Round(Offset1, 3) & Left_right, 1, 0.2, 0.5, 5, 5)

                                                    new_leader.Layer = No_plot


                                                    Dim Linie1 As New Line(New Point3d(Block1.Position.X, Block1.Position.Y, 0), Point_on_poly2d)
                                                    Linie1.Layer = No_plot

                                                    BTrecord.AppendEntity(Linie1)

                                                    Trans1.AddNewlyCreatedDBObject(Linie1, True)

                                                End If

                                            End If


                                            If CheckBox_scan_points.Checked = True Then
                                                If TypeOf Ent_Object Is DBPoint Then
                                                    Dim DBPoint1 As DBPoint = Ent_Object
                                                    Dim Point_on_poly As New Point3d
                                                    Dim Point_on_poly3d As New Point3d
                                                    Dim Point_on_poly2d As New Point3d

                                                    Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(DBPoint1.Position.X, DBPoint1.Position.Y, 0), Vector3d.ZAxis, True)

                                                    Dim Offset1 As Double = New Point3d(DBPoint1.Position.X, DBPoint1.Position.Y, 0).GetVectorTo(Point_on_poly2d).Length

                                                    If Offset1 <= CDbl(TextBox_buffer.Text) Then

                                                        Data_Table_information.Rows.Add()
                                                        Dim ChainageGrid0 As Double

                                                        If Not Poly3d = Nothing Then
                                                            Dim Param2d As Double = Poly2D.GetParameterAtPoint(Point_on_poly2d)
                                                            Point_on_poly3d = Poly3d.GetPointAtParameter(Param2d)
                                                            Point_on_poly = Point_on_poly3d
                                                            ChainageGrid0 = Poly3d.GetDistAtPoint(Point_on_poly3d)
                                                        Else
                                                            Point_on_poly = Point_on_poly2d
                                                            ChainageGrid0 = Poly2D.GetDistAtPoint(Point_on_poly2d)
                                                        End If



                                                        Data_Table_information.Rows(Index_row).Item(Col_x) = Round(Point_on_poly.X, 3)
                                                        Data_Table_information.Rows(Index_row).Item(Col_y) = Round(Point_on_poly.Y, 3)
                                                        Data_Table_information.Rows(Index_row).Item(Col_z) = Round(Point_on_poly.Z, 3)
                                                        Data_Table_information.Rows(Index_row).Item(Col_z_point) = Round(DBPoint1.Position.Z, 3)
                                                        Data_Table_information.Rows(Index_row).Item(Col_station) = Round(ChainageGrid0, nrdec)
                                                        Data_Table_information.Rows(Index_row).Item(Col_station_eq) = Round(ChainageGrid0 + Get_equation_value(ChainageGrid0), nrdec)
                                                        Data_Table_information.Rows(Index_row).Item(Col_offset) = Round(Offset1, nrdec)
                                                        Data_Table_information.Rows(Index_row).Item(Col_block_name) = "POINT"

                                                        Dim Left_right As String = Angle_left_right(Poly2D, New Point3d(DBPoint1.Position.X, DBPoint1.Position.Y, 0), 1)
                                                        Data_Table_information.Rows(Index_row).Item(Col_left_right) = Left_right



                                                        Dim Valoare As String = DBPoint1.Layer
                                                        Dim Valoare1 As String = ""
                                                        If IsNothing(Data_table_Layers) = False Then
                                                            If Data_table_Layers.Rows.Count > 0 Then
                                                                For i = 0 To Data_table_Layers.Rows.Count - 1
                                                                    If IsDBNull(Data_table_Layers.Rows(i).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(i).Item(Col_description)) = False Then
                                                                        If Data_table_Layers.Rows(i).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                            Valoare1 = Data_table_Layers.Rows(i).Item(Col_description)
                                                                            Exit For
                                                                        End If

                                                                    End If
                                                                Next
                                                            End If
                                                        End If

                                                        Data_Table_information.Rows(Index_row).Item(Col_layer_name) = DBPoint1.Layer
                                                        If Not Valoare1 = "" Then
                                                            Data_Table_information.Rows(Index_row).Item(Col_layer_description) = Valoare1
                                                        End If
                                                        Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), DBPoint1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                            If IsNothing(Records1) = False Then
                                                                If Records1.Count > 0 Then
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

                                                                            If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                                Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                            End If
                                                                            If Not Replace(Valoare_field, " ", "") = "" Then
                                                                                Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                            End If
                                                                        Next
                                                                    Next
                                                                End If
                                                            End If
                                                        End Using

                                                        Index_row = Index_row + 1

                                                        Dim new_leader As New MLeader


                                                        Dim Station As String
                                                        Dim Station_EQ As String
                                                        If CheckBox_US_station.Checked = False Then
                                                            Station = Get_chainage_from_double(ChainageGrid0, 3)
                                                            Station_EQ = Get_chainage_from_double(ChainageGrid0 + Get_equation_value(ChainageGrid0), 3)
                                                        Else
                                                            Station = Get_chainage_feet_from_double(ChainageGrid0, 0)
                                                            Station_EQ = Get_chainage_feet_from_double(ChainageGrid0 + Get_equation_value(ChainageGrid0), 0)
                                                        End If
                                                        Dim eXTRA1 As String = ""
                                                        If Not Get_equation_value(ChainageGrid0) = 0 Then
                                                            eXTRA1 = "St_eq = " & Station_EQ & vbCrLf
                                                        End If


                                                        new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly,
                                                                                                           "East = " & Round(Point_on_poly.X, 3) & vbCrLf &
                                                                                                           "North = " & Round(Point_on_poly.Y, 3) & vbCrLf &
                                                                                                           "Elev = " & Round(Point_on_poly.Z, 3) & vbCrLf &
                                                                                                           "Station = " & Station & vbCrLf & eXTRA1 &
                                                                                                           "Offset = " & Round(Offset1, 3) & Left_right, 1, 0.2, 0.5, 5, 5)


                                                        new_leader.Layer = No_plot

                                                        Dim Linie1 As New Line(New Point3d(DBPoint1.Position.X, DBPoint1.Position.Y, 0), Point_on_poly)
                                                        Linie1.Layer = No_plot
                                                        BTrecord.AppendEntity(Linie1)
                                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)


                                                    End If
                                                End If
                                            End If






                                        End If
                                    End If


                                Next








                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                If Data_Table_information.Rows.Count > 0 Then
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                    W1.Cells.NumberFormat = "@"

                    Dim maxRows As Integer = Data_Table_information.Rows.Count
                    Dim maxCols As Integer = Data_Table_information.Columns.Count

                    Dim range As Microsoft.Office.Interop.Excel.Range = W1.Range(W1.Cells(2, 1), W1.Cells(maxRows + 1, maxCols))

                    Dim values(maxRows, maxCols) As Object

                    For row = 0 To maxRows - 1
                        For col = 0 To maxCols - 1
                            If IsDBNull(Data_Table_information.Rows(row).Item(col)) = False Then
                                values(row, col) = Data_Table_information.Rows(row).Item(col)
                            End If
                        Next
                    Next

                    range.Value2 = values

                    For i = 0 To Data_Table_information.Columns.Count - 1
                        W1.Cells(1, i + 1).value2 = Data_Table_information.Columns(i).ColumnName
                    Next
                End If





                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Private Sub Button_LOAD_LAYER_NAMES_Click(sender As Object, e As EventArgs) Handles Button_LOAD_LAYER_NAMES.Click
        Try
            If TextBox_layer_name.Text = "" Then
                MsgBox("Please specify the Layer Name COLUMN!")
                Exit Sub
            End If
            If TextBox_layer_description.Text = "" Then
                MsgBox("Please specify the Layer Description COLUMN!")
                Exit Sub
            End If

            If IsNumeric(TextBox_start_row.Text) = False Then
                MsgBox("Please specify the start row!")
                Exit Sub
            End If
            If IsNumeric(TextBox_end_row.Text) = False Then
                MsgBox("Please specify the end row!")
                Exit Sub
            End If

            Dim Start1 As Integer = Abs(CInt(TextBox_start_row.Text))
            Dim End1 As Integer = Abs(CInt(TextBox_end_row.Text))

            If Start1 > End1 Then
                Dim T As Integer = Start1
                Start1 = End1
                End1 = T
            End If
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Data_table_Layers = New System.Data.DataTable
            Data_table_Layers.Columns.Add("NAME", GetType(String))
            Data_table_Layers.Columns.Add("DESCRIPTION", GetType(String))
            Dim Index1 As Integer = 0
            Button_LOAD_LAYER_NAMES.Visible = False
            For i = Start1 To End1
                Dim Name As String = W1.Range(TextBox_layer_name.Text.ToUpper & i).Value2
                Dim Description As String = W1.Range(TextBox_layer_description.Text.ToUpper & i).Value2
                If Len(Name) > 0 And Len(Description) > 0 Then
                    Data_table_Layers.Rows.Add()
                    Data_table_Layers.Rows(Index1).Item("NAME") = Name
                    Data_table_Layers.Rows(Index1).Item("DESCRIPTION") = Description
                    Index1 = Index1 + 1
                End If

            Next



            Button_LOAD_LAYER_NAMES.Visible = True

        Catch ex As Exception
            Button_LOAD_LAYER_NAMES.Visible = True

            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_calculate_int_Click(sender As Object, e As EventArgs) Handles Button_calculate_int.Click
        If Freeze_operations = False Then
            Try

                Dim Col_station As String = "STATION"
                Dim Col_station_eq As String = "STATION_EQ"

                Dim Data_Table_information As New System.Data.DataTable
                Data_Table_information.Columns.Add(Col_station, GetType(Double))
                Data_Table_information.Columns.Add(Col_station_eq, GetType(Double))
                Data_Table_information.Columns.Add("DESCRIPTION", GetType(String))

                Data_Table_information.Columns.Add("X", GetType(Double))
                Data_Table_information.Columns.Add("Y", GetType(Double))
                Data_Table_information.Columns.Add("Z", GetType(Double))
                If CheckBox_Line_direction.Checked = True Then
                    Data_Table_information.Columns.Add("From Left To Right", GetType(String))
                End If

                If CheckBox_output_layers.Checked = True Then
                    Data_Table_information.Columns.Add("LAYER_NAME", GetType(String))
                    Data_Table_information.Columns.Add("LAYER_DESCRIPTION", GetType(String))
                End If
                If CheckBox_no_CSF.Checked = False Then
                    Data_Table_information.Columns.Add("STATION_CSF", GetType(Double))
                End If
                If CheckBox_select_multiple_CL.Checked = True Then
                    Data_Table_information.Columns.Add("CL_NO", GetType(Double))
                End If

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    If CheckBox_select_multiple_CL.Checked = True Then
                        Object_Prompt.MessageForAdding = vbLf & "Select one CL (all objects from the CL layer will be treated as CL):"
                    Else
                        Object_Prompt.MessageForAdding = vbLf & "Select centerline:"
                    End If

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)
                                Dim Col_ObjId_CL As New ObjectIdCollection

                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)



                                Dim Index_row As Double = 0

                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection

                                If TypeOf Ent1 Is Polyline3d Then
                                    Col_ObjId_CL.Add(Obj1.ObjectId)
                                End If

                                If TypeOf Ent1 Is Polyline Then
                                    Col_ObjId_CL.Add(Obj1.ObjectId)
                                End If

                                If Col_ObjId_CL.Count = 0 Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                If CheckBox_no_CSF.Checked = False Then
                                    For Each ObjID In BTrecord
                                        If CheckBox_no_CSF.Checked = False Then
                                            Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                            If TypeOf DBobject Is DBText Then
                                                Dim DBText1 As DBText = DBobject
                                                If DBText1.TextString.Contains("+") = True Then
                                                    ChainageCSF_colection.Add(DBText1)
                                                End If
                                                If DBText1.TextString.Contains("CSF") = True Then
                                                    CSF_colection.Add(DBText1)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If


                                If CheckBox_select_multiple_CL.Checked = True Then
                                    For Each ObjID In BTrecord
                                        If CheckBox_select_multiple_CL.Checked = True Then
                                            Dim ENT_CL As Entity = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                                            If TypeOf ENT_CL Is Polyline Or TypeOf ENT_CL Is Polyline3d Then
                                                If ENT_CL.Layer = Ent1.Layer Then
                                                    If Col_ObjId_CL.Contains(ObjID) = False Then
                                                        Col_ObjId_CL.Add(ObjID)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                End If

                                'If CSF_colection.Count > 0 And ChainageCSF_colection.Count > 0 Then


                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables

                                Dim nrdec As Integer = 3
                                If CheckBox_zero_decimals.Checked = True Then nrdec = 0

                                For Each ObjID In BTrecord
                                    Dim Ent_intersection As Entity = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)

                                    If Not Ent_intersection = Nothing Then
                                        If TypeOf Ent_intersection Is Curve And Col_ObjId_CL.Contains(Ent_intersection.ObjectId) = False Then


                                            Dim Poly_int As Curve = Ent_intersection

                                            For i = 0 To Col_ObjId_CL.Count - 1
                                                Dim Ent_CL As Entity = Trans1.GetObject(Col_ObjId_CL(i), OpenMode.ForRead)

                                                Dim Poly2D As New Polyline
                                                Dim Poly3D As Polyline3d

                                                If TypeOf Ent_CL Is Polyline3d Then
                                                    Poly3D = Ent_CL
                                                    Dim Index2d As Double = 0
                                                    For Each ObjId1 As ObjectId In Poly3D
                                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                                        Index2d = Index2d + 1
                                                    Next
                                                End If


                                                If TypeOf Ent_CL Is Polyline Then
                                                    Poly2D = Ent_CL.Clone
                                                End If

                                                Poly2D.Elevation = 0



                                                If TypeOf Poly_int Is Polyline Then
                                                    Dim Poly2 As Polyline = Poly_int
                                                    Poly2D.Elevation = Poly2.Elevation

                                                End If
                                                If TypeOf Poly_int Is Line Then
                                                    Dim Line2 As Line = Poly_int
                                                    Poly2D.Elevation = Line2.StartPoint.Z
                                                End If

                                                Dim point_dir As Point3d = New Point3d()
                                                point_dir = Poly_int.EndPoint
                                                Dim cn As String = "From Left To Right"
                                                Dim cn_value As String


                                                If CheckBox_Line_direction.Checked = True Then

                                                    Dim Left_right As String = Angle_left_right(Poly2D, New Point3d(point_dir.X, point_dir.Y, 0), 1)
                                                    If Left_right = " LT." Then
                                                        cn_value = "no"
                                                    Else
                                                        cn_value = "yes"

                                                    End If

                                                End If


                                                Dim Col_int As New Point3dCollection

                                                Col_int = Intersect_on_both_operands(Poly_int, Poly2D)



                                                If Col_int.Count > 0 Then

                                                    For index = 0 To Col_int.Count - 1
                                                        Dim Point_on_poly2d As New Point3d
                                                        Dim Point_on_poly As New Point3d
                                                        Dim ChainageGrid0 As Double = 0
                                                        Dim ChainageCSF0 As Double = 0

                                                        Point_on_poly2d = Poly2D.GetClosestPointTo(Col_int(index), Vector3d.ZAxis, True)
                                                        Dim Param2d As Double = Poly2D.GetParameterAtPoint(Point_on_poly2d)



                                                        If TypeOf Ent_CL Is Polyline3d Then
                                                            Point_on_poly = Poly3D.GetPointAtParameter(Param2d)
                                                            ChainageGrid0 = Poly3D.GetDistAtPoint(Point_on_poly)
                                                            If CheckBox_no_CSF.Checked = False Then
                                                                ChainageCSF0 = Get_chainage_with_CSF_from_dbtext(Poly3D, Point_on_poly, CSF_colection, ChainageCSF_colection)
                                                            End If

                                                        End If


                                                        If TypeOf Ent_CL Is Polyline Then
                                                            Point_on_poly = Point_on_poly2d
                                                            ChainageGrid0 = Poly2D.GetDistAtPoint(Point_on_poly)
                                                            If CheckBox_no_CSF.Checked = False Then
                                                                ChainageCSF0 = Get_chainage_with_CSF_from_dbtext(Poly2D, Point_on_poly, CSF_colection, ChainageCSF_colection)
                                                            End If
                                                        End If


                                                        Data_Table_information.Rows.Add()
                                                        Data_Table_information.Rows(Index_row).Item("X") = Round(Point_on_poly.X, 3)
                                                        Data_Table_information.Rows(Index_row).Item("Y") = Round(Point_on_poly.Y, 3)
                                                        Data_Table_information.Rows(Index_row).Item("Z") = Round(Point_on_poly.Z, 3)
                                                        If CheckBox_Line_direction.Checked = True Then Data_Table_information.Rows(Index_row).Item(cn) = cn_value

                                                        Data_Table_information.Rows(Index_row).Item(Col_station) = Round(ChainageGrid0, nrdec)
                                                        Data_Table_information.Rows(Index_row).Item(Col_station_eq) = Round(ChainageGrid0 + Get_equation_value(ChainageGrid0), nrdec)



                                                        If CheckBox_select_multiple_CL.Checked = True Then
                                                            Data_Table_information.Rows(Index_row).Item("CL_NO") = i + 1
                                                        End If


                                                        If CheckBox_no_CSF.Checked = False Then
                                                            If CheckBox_use_equation.Checked = False Then
                                                                Data_Table_information.Rows(Index_row).Item("STATION_CSF") = Round(ChainageCSF0, nrdec)
                                                            Else
                                                                If IsNothing(Data_table_station_equation) = False Then
                                                                    If Data_table_station_equation.Rows.Count > 0 Then
                                                                        Data_Table_information.Rows(Index_row).Item("STATION_CSF") = Round(ChainageGrid0, nrdec) + Get_equation_value(ChainageGrid0)
                                                                    Else
                                                                        Data_Table_information.Rows(Index_row).Item("STATION_CSF") = Round(ChainageGrid0, nrdec)
                                                                    End If
                                                                Else
                                                                    Data_Table_information.Rows(Index_row).Item("STATION_CSF") = Round(ChainageGrid0, nrdec)
                                                                End If
                                                            End If


                                                        End If


                                                        If CheckBox_output_layers.Checked = True Then
                                                            Dim Valoare As String = Poly_int.Layer
                                                            Dim Valoare1 As String = ""
                                                            If IsNothing(Data_table_Layers) = False Then
                                                                If Data_table_Layers.Rows.Count > 0 Then
                                                                    For J = 0 To Data_table_Layers.Rows.Count - 1
                                                                        If IsDBNull(Data_table_Layers.Rows(J).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(J).Item("DESCRIPTION")) = False Then
                                                                            If Data_table_Layers.Rows(J).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                                Valoare1 = Data_table_Layers.Rows(J).Item("DESCRIPTION")
                                                                                Exit For
                                                                            End If

                                                                        End If
                                                                    Next
                                                                End If
                                                            End If
                                                            Data_Table_information.Rows(Index_row).Item("LAYER_NAME") = Poly_int.Layer
                                                            If Not Valoare1 = "" Then
                                                                Data_Table_information.Rows(Index_row).Item("LAYER_DESCRIPTION") = Valoare1
                                                            End If
                                                        End If

                                                        If CheckBox_Object_data.Checked = True Then
                                                            Dim Id1 As ObjectId = Poly_int.ObjectId
                                                            Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                                If IsNothing(Records1) = False Then
                                                                    If Records1.Count > 0 Then
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

                                                                                If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                                    Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                                End If
                                                                                If Not Replace(Valoare_field, " ", "") = "" Then
                                                                                    Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                                End If
                                                                            Next
                                                                        Next
                                                                    End If
                                                                End If
                                                            End Using

                                                        End If

                                                        Index_row = Index_row + 1

                                                        Dim new_leader As New MLeader

                                                        Dim Station As String
                                                        Dim StationCSF As String

                                                        Dim Station_EQ As String
                                                        If CheckBox_US_station.Checked = False Then
                                                            Station = Get_chainage_from_double(ChainageGrid0, 3)
                                                            Station_EQ = Get_chainage_from_double(ChainageGrid0 + Get_equation_value(ChainageGrid0), 3)
                                                            StationCSF = Get_chainage_from_double(ChainageCSF0, 3)
                                                        Else
                                                            Station = Get_chainage_feet_from_double(ChainageGrid0, 0)
                                                            Station_EQ = Get_chainage_feet_from_double(ChainageGrid0 + Get_equation_value(ChainageGrid0), 0)
                                                            StationCSF = Get_chainage_from_double(ChainageCSF0, 3)
                                                        End If
                                                        Dim eXTRA1 As String = ""
                                                        If Not Get_equation_value(ChainageGrid0) = 0 Then
                                                            eXTRA1 = "St_eq = " & Station_EQ & vbCrLf
                                                        End If



                                                        If CheckBox_no_CSF.Checked = True Then

                                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly,
                                                                                                               "East = " & Round(Point_on_poly.X, 3) & vbCrLf &
                                                                                                               "North = " & Round(Point_on_poly.Y, 3) & vbCrLf &
                                                                                                               "Elev = " & Round(Point_on_poly.Z, 3) & vbCrLf & eXTRA1 &
                                                                                                               "Station Grid = " & Station, 1, 0.2, 0.5, 5, 5)
                                                            new_leader.Layer = "NO PLOT"

                                                        Else

                                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly,
                                                               "East = " & Round(Point_on_poly.X, 3) & vbCrLf &
                                                               "North = " & Round(Point_on_poly.Y, 3) & vbCrLf &
                                                               "Elev = " & Round(Point_on_poly.Z, 3) & vbCrLf &
                                                               "Grid Station = " & Station & vbCrLf &
                                                               "CSF Station = " & StationCSF _
                                                               , 1, 0.2, 0.5, 5, 5)
                                                            new_leader.Layer = "NO PLOT"
                                                        End If
                                                    Next
                                                End If



                                            Next


                                        End If
                                    End If
                                Next






                                Editor1.Regen()
                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                Transfer_datatable_to_new_excel_spreadsheet(Data_Table_information)


                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button_calculate_start_end_Click(sender As Object, e As EventArgs) Handles Button_calc_start_end.Click
        If Freeze_operations = False Then
            Try

                Dim Col_station1 As String = "STATION1"
                Dim Col_station_eq1 As String = "STATION_EQ1"
                Dim Col_station2 As String = "STATION2"
                Dim Col_station_eq2 As String = "STATION_EQ2"
                Dim Data_Table_information As New System.Data.DataTable
                Data_Table_information.Columns.Add(Col_station1, GetType(Double))
                Data_Table_information.Columns.Add(Col_station2, GetType(Double))
                Data_Table_information.Columns.Add(Col_station_eq1, GetType(Double))
                Data_Table_information.Columns.Add(Col_station_eq2, GetType(Double))
                Data_Table_information.Columns.Add("DESCRIPTION", GetType(String))

                Data_Table_information.Columns.Add("X", GetType(Double))
                Data_Table_information.Columns.Add("Y", GetType(Double))
                Data_Table_information.Columns.Add("Z", GetType(Double))

                If CheckBox_output_layers.Checked = True Then
                    Data_Table_information.Columns.Add("LAYER_NAME", GetType(String))
                    Data_Table_information.Columns.Add("LAYER_DESCRIPTION", GetType(String))
                End If
                If CheckBox_no_CSF.Checked = False Then
                    Data_Table_information.Columns.Add("STATION_CSF", GetType(Double))
                End If
                If CheckBox_select_multiple_CL.Checked = True Then
                    Data_Table_information.Columns.Add("CL_NO", GetType(Double))
                End If

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    If CheckBox_select_multiple_CL.Checked = True Then
                        Object_Prompt.MessageForAdding = vbLf & "Select one CL (all objects from the CL layer will be treated as CL):"
                    Else
                        Object_Prompt.MessageForAdding = vbLf & "Select centerline:"
                    End If

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)
                                Dim Col_ObjId_CL As New ObjectIdCollection

                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)



                                Dim Index_row As Double = 0

                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection

                                If TypeOf Ent1 Is Polyline3d Then
                                    Col_ObjId_CL.Add(Obj1.ObjectId)
                                End If

                                If TypeOf Ent1 Is Polyline Then
                                    Col_ObjId_CL.Add(Obj1.ObjectId)
                                End If

                                If Col_ObjId_CL.Count = 0 Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                If CheckBox_no_CSF.Checked = False Then
                                    For Each ObjID In BTrecord
                                        If CheckBox_no_CSF.Checked = False Then
                                            Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                            If TypeOf DBobject Is DBText Then
                                                Dim DBText1 As DBText = DBobject
                                                If DBText1.TextString.Contains("+") = True Then
                                                    ChainageCSF_colection.Add(DBText1)
                                                End If
                                                If DBText1.TextString.Contains("CSF") = True Then
                                                    CSF_colection.Add(DBText1)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If


                                If CheckBox_select_multiple_CL.Checked = True Then
                                    For Each ObjID In BTrecord
                                        If CheckBox_select_multiple_CL.Checked = True Then
                                            Dim ENT_CL As Entity = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                                            If TypeOf ENT_CL Is Polyline Or TypeOf ENT_CL Is Polyline3d Then
                                                If ENT_CL.Layer = Ent1.Layer Then
                                                    If Col_ObjId_CL.Contains(ObjID) = False Then
                                                        Col_ObjId_CL.Add(ObjID)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                End If

                                'If CSF_colection.Count > 0 And ChainageCSF_colection.Count > 0 Then


                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables

                                Dim nrdec As Integer = 3
                                If CheckBox_zero_decimals.Checked = True Then nrdec = 0

                                For Each ObjID In BTrecord
                                    Dim entity_on_top As Curve = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Curve)

                                    If Not entity_on_top = Nothing Then
                                        If Col_ObjId_CL.Contains(entity_on_top.ObjectId) = False Then


                                            Dim Poly_int As Curve = entity_on_top

                                            For i = 0 To Col_ObjId_CL.Count - 1
                                                Dim Ent_CL As Entity = Trans1.GetObject(Col_ObjId_CL(i), OpenMode.ForRead)

                                                Dim Poly2D As New Polyline
                                                Dim Poly3D As Polyline3d

                                                If TypeOf Ent_CL Is Polyline3d Then
                                                    Poly3D = Ent_CL
                                                    Dim Index2d As Double = 0
                                                    For Each ObjId1 As ObjectId In Poly3D
                                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                                        Index2d = Index2d + 1
                                                    Next
                                                End If


                                                If TypeOf Ent_CL Is Polyline Then
                                                    Poly2D = Ent_CL.Clone
                                                End If

                                                Poly2D.Elevation = 0

                                                If TypeOf Poly_int Is Polyline Then
                                                    Dim Poly2 As Polyline = Poly_int
                                                    Poly2D.Elevation = Poly2.Elevation
                                                End If
                                                If TypeOf Poly_int Is Line Then
                                                    Dim Line2 As Line = Poly_int
                                                    Poly2D.Elevation = Line2.StartPoint.Z
                                                End If







                                                Dim Point_on_poly2d1 As New Point3d
                                                Dim Point_on_poly2d2 As New Point3d
                                                Dim Point_on_poly1 As New Point3d
                                                Dim Point_on_poly2 As New Point3d
                                                Dim Chainage1 As Double = 0
                                                Dim Chainage2 As Double = 0

                                                Dim Start1 As Point3d = Poly_int.StartPoint
                                                Dim End1 As Point3d = Poly_int.EndPoint

                                                Point_on_poly2d1 = Poly2D.GetClosestPointTo(Start1, Vector3d.ZAxis, True)
                                                Dim Param2d1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2d1)
                                                Point_on_poly2d2 = Poly2D.GetClosestPointTo(End1, Vector3d.ZAxis, True)
                                                Dim Param2d2 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2d2)

                                                If TypeOf Ent_CL Is Polyline3d Then
                                                    Point_on_poly1 = Poly3D.GetPointAtParameter(Param2d1)
                                                    Chainage1 = Poly3D.GetDistAtPoint(Point_on_poly1)
                                                    Point_on_poly2 = Poly3D.GetPointAtParameter(Param2d2)
                                                    Chainage2 = Poly3D.GetDistAtPoint(Point_on_poly2)
                                                End If


                                                If TypeOf Ent_CL Is Polyline Then
                                                    Point_on_poly1 = Point_on_poly2d1
                                                    Chainage1 = Poly2D.GetDistAtPoint(Point_on_poly1)
                                                    Point_on_poly2 = Point_on_poly2d2
                                                    Chainage2 = Poly2D.GetDistAtPoint(Point_on_poly2)
                                                End If


                                                Data_Table_information.Rows.Add()
                                                Data_Table_information.Rows(Index_row).Item("X") = Round(Point_on_poly1.X, 3)
                                                Data_Table_information.Rows(Index_row).Item("Y") = Round(Point_on_poly1.Y, 3)
                                                Data_Table_information.Rows(Index_row).Item("Z") = Round(Point_on_poly1.Z, 3)
                                                Data_Table_information.Rows(Index_row).Item(Col_station1) = Round(Chainage1, nrdec)
                                                Data_Table_information.Rows(Index_row).Item(Col_station_eq1) = Round(Chainage1 + Get_equation_value(Chainage1), nrdec)
                                                Data_Table_information.Rows(Index_row).Item(Col_station2) = Round(Chainage2, nrdec)
                                                Data_Table_information.Rows(Index_row).Item(Col_station_eq2) = Round(Chainage2 + Get_equation_value(Chainage2), nrdec)


                                                If CheckBox_select_multiple_CL.Checked = True Then
                                                    Data_Table_information.Rows(Index_row).Item("CL_NO") = i + 1
                                                End If





                                                If CheckBox_output_layers.Checked = True Then
                                                    Dim Valoare As String = Poly_int.Layer
                                                    Dim Valoare1 As String = ""
                                                    If IsNothing(Data_table_Layers) = False Then
                                                        If Data_table_Layers.Rows.Count > 0 Then
                                                            For J = 0 To Data_table_Layers.Rows.Count - 1
                                                                If IsDBNull(Data_table_Layers.Rows(J).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(J).Item("DESCRIPTION")) = False Then
                                                                    If Data_table_Layers.Rows(J).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                        Valoare1 = Data_table_Layers.Rows(J).Item("DESCRIPTION")
                                                                        Exit For
                                                                    End If

                                                                End If
                                                            Next
                                                        End If
                                                    End If
                                                    Data_Table_information.Rows(Index_row).Item("LAYER_NAME") = Poly_int.Layer
                                                    If Not Valoare1 = "" Then
                                                        Data_Table_information.Rows(Index_row).Item("LAYER_DESCRIPTION") = Valoare1
                                                    End If
                                                End If

                                                If CheckBox_Object_data.Checked = True Then
                                                    Dim Id1 As ObjectId = Poly_int.ObjectId
                                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                        If IsNothing(Records1) = False Then
                                                            If Records1.Count > 0 Then
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

                                                                        If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                            Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                        End If
                                                                        If Not Replace(Valoare_field, " ", "") = "" Then
                                                                            Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If
                                                        End If
                                                    End Using

                                                End If

                                                Index_row = Index_row + 1

                                                Dim new_leader1 As New MLeader
                                                Dim Station1 As String
                                                Dim Station_EQ1 As String
                                                If CheckBox_US_station.Checked = False Then
                                                    Station1 = Get_chainage_from_double(Chainage1, 3)
                                                    Station_EQ1 = Get_chainage_from_double(Chainage1 + Get_equation_value(Chainage1), 3)
                                                Else
                                                    Station1 = Get_chainage_feet_from_double(Chainage1, 0)
                                                    Station_EQ1 = Get_chainage_feet_from_double(Chainage1 + Get_equation_value(Chainage1), 0)
                                                End If
                                                Dim extra1 As String = ""
                                                If Not Get_equation_value(Chainage1) = 0 Then
                                                    extra1 = vbCrLf & "St_eq = " & Station_EQ1
                                                End If
                                                new_leader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly1,
                                                       "East = " & Round(Point_on_poly1.X, 3) & vbCrLf &
                                                       "North = " & Round(Point_on_poly1.Y, 3) & vbCrLf &
                                                       "Elev = " & Round(Point_on_poly1.Z, 3) & vbCrLf &
                                                       "Grid Station = " & Station1 & extra1, 1, 0.2, 0.5, 5, 5)
                                                new_leader1.Layer = "NO PLOT"

                                                Dim new_leader2 As New MLeader
                                                Dim Station2 As String
                                                Dim Station_EQ2 As String
                                                If CheckBox_US_station.Checked = False Then
                                                    Station2 = Get_chainage_from_double(Chainage2, 3)
                                                    Station_EQ2 = Get_chainage_from_double(Chainage2 + Get_equation_value(Chainage2), 3)
                                                Else
                                                    Station2 = Get_chainage_feet_from_double(Chainage2, 0)
                                                    Station_EQ2 = Get_chainage_feet_from_double(Chainage2 + Get_equation_value(Chainage2), 0)
                                                End If
                                                Dim extra2 As String = ""
                                                If Not Get_equation_value(Chainage2) = 0 Then
                                                    extra2 = vbCrLf & "St_eq = " & Station_EQ2
                                                End If
                                                new_leader2 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly2,
                                                       "East = " & Round(Point_on_poly2.X, 3) & vbCrLf &
                                                       "North = " & Round(Point_on_poly2.Y, 3) & vbCrLf &
                                                       "Elev = " & Round(Point_on_poly2.Z, 3) & vbCrLf &
                                                       "Grid Station = " & Station2 & extra2, 2, 0.2, 0.5, 5, 5)
                                                new_leader2.Layer = "NO PLOT"



                                            Next


                                        End If
                                    End If
                                Next






                                Editor1.Regen()
                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                Transfer_datatable_to_new_excel_spreadsheet(Data_Table_information)


                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button_Scan_segments_Click(sender As Object, e As EventArgs) Handles Button_Scan_segments.Click
        If Freeze_operations = False Then

            Try



                If IsNumeric(TextBox_buffer.Text) = False Then
                    MsgBox("Please specify the Buffer size!")
                    Exit Sub
                End If


                Dim Data_Table_information As New System.Data.DataTable
                Data_Table_information.Columns.Add("STATION", GetType(Double))
                Data_Table_information.Columns.Add("DESCRIPTION", GetType(String))



                Data_Table_information.Columns.Add("X", GetType(Double))
                Data_Table_information.Columns.Add("Y", GetType(Double))
                Data_Table_information.Columns.Add("Z", GetType(Double))
                Data_Table_information.Columns.Add("SEGMENT_NUMBER", GetType(Integer))

                Data_Table_information.Columns.Add("LAYER_NAME", GetType(String))
                Data_Table_information.Columns.Add("LAYER_DESCRIPTION", GetType(String))


                Data_Table_information.Columns.Add("OFFSET", GetType(Double))
                Data_Table_information.Columns.Add("BLOCK_NAME", GetType(String))
                Data_Table_information.Columns.Add("LEFT_RIGHT", GetType(String))



                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select centerline:"

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                Dim Poly3d As Polyline3d = Nothing
                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)





                                Dim Index_row As Double = 0

                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection


                                Dim Poly2D As New Polyline


                                If TypeOf Ent1 Is Polyline3d Then
                                    Poly3d = Ent1
                                    Dim Index2d As Double = 0
                                    For Each ObjId As ObjectId In Poly3d
                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                        Index2d = Index2d + 1
                                    Next
                                End If


                                If TypeOf Ent1 Is Polyline Then
                                    Poly2D = Ent1.Clone
                                End If

                                If Poly2D = Nothing Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Poly2D.Elevation = 0




                                Dim Data_table_Columns_object_data As New System.Data.DataTable
                                Data_table_Columns_object_data.Columns.Add("NAME", GetType(String))
                                Data_table_Columns_object_data.Columns.Add("NR", GetType(Integer))
                                Dim nrdec As Integer = 3
                                If CheckBox_zero_decimals.Checked = True Then nrdec = 0
                                Dim Segment_No As Integer = 1

                                Dim Line_new() As Autodesk.AutoCAD.DatabaseServices.Line
                                Dim Index_l As Integer = 1

                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables


                                For Each ObjID In BTrecord

                                    Dim Ent_Object As Entity = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                                    If Not Ent_Object = Nothing Then
                                        If Not ObjID = Ent1.ObjectId Then

                                            If TypeOf (Ent_Object) Is Polyline Or TypeOf (Ent_Object) Is Line Then
                                                Dim Object_ID_OD As ObjectId

                                                Dim Line_OBJ As Autodesk.AutoCAD.DatabaseServices.Line = TryCast(Ent_Object, Autodesk.AutoCAD.DatabaseServices.Line)

                                                Dim Pline_OBJ As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(Ent_Object, Autodesk.AutoCAD.DatabaseServices.Polyline)

                                                If IsNothing(Line_OBJ) = False Then
                                                    Pline_OBJ = New Polyline
                                                    Pline_OBJ.AddVertexAt(0, New Point2d(Line_OBJ.StartPoint.X, Line_OBJ.StartPoint.Y), 0, 0, 0)
                                                    Pline_OBJ.AddVertexAt(1, New Point2d(Line_OBJ.EndPoint.X, Line_OBJ.EndPoint.Y), 0, 0, 0)
                                                    Object_ID_OD = Line_OBJ.ObjectId
                                                Else
                                                    Object_ID_OD = Pline_OBJ.ObjectId
                                                End If


                                                Dim Data_Table_object As New System.Data.DataTable
                                                Data_Table_object.Columns.Add("X", GetType(Double))
                                                Data_Table_object.Columns.Add("Y", GetType(Double))
                                                Data_Table_object.Columns.Add("Z", GetType(Double))
                                                Data_Table_object.Columns.Add("SEGMENT_NUMBER", GetType(Integer))
                                                Data_Table_object.Columns.Add("STATION", GetType(Double))

                                                Data_Table_object.Columns.Add("LAYER_NAME", GetType(String))
                                                Data_Table_object.Columns.Add("LAYER_DESCRIPTION", GetType(String))


                                                Data_Table_object.Columns.Add("OFFSET", GetType(Double))
                                                Data_Table_object.Columns.Add("BLOCK_NAME", GetType(String))
                                                Data_Table_object.Columns.Add("LEFT_RIGHT", GetType(String))
                                                Data_Table_object.Columns.Add("X1", GetType(Double))
                                                Data_Table_object.Columns.Add("Y1", GetType(Double))
                                                Data_Table_object.Columns.Add("Z1", GetType(Double))
                                                Dim Index__obj As Double = 0
                                                For j = 0 To Pline_OBJ.NumberOfVertices - 1


                                                    Dim Point_on_poly As New Point3d
                                                    Dim Point_on_poly3d As New Point3d
                                                    Dim Point_on_poly2d As New Point3d


                                                    Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0), Vector3d.ZAxis, True)

                                                    Dim Offset1 As Double = New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0).GetVectorTo(Point_on_poly2d).Length





                                                    If Offset1 <= CDbl(TextBox_buffer.Text) Then

                                                        Dim ChainageGrid0 As Double

                                                        If Not Poly3d = Nothing Then
                                                            Dim Param2d As Double = Poly2D.GetParameterAtPoint(Point_on_poly2d)
                                                            Point_on_poly3d = Poly3d.GetPointAtParameter(Param2d)
                                                            Point_on_poly = Point_on_poly3d
                                                            ChainageGrid0 = Poly3d.GetDistAtPoint(Point_on_poly3d)
                                                        Else
                                                            Point_on_poly = Point_on_poly2d
                                                            ChainageGrid0 = Poly2D.GetDistAtPoint(Point_on_poly2d)
                                                        End If




                                                        Data_Table_object.Rows.Add()
                                                        Data_Table_object.Rows(Index__obj).Item("X1") = Round(Pline_OBJ.GetPoint3dAt(j).X, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Y1") = Round(Pline_OBJ.GetPoint3dAt(j).Y, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Z1") = 0

                                                        Data_Table_object.Rows(Index__obj).Item("X") = Round(Point_on_poly.X, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Y") = Round(Point_on_poly.Y, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Z") = 0
                                                        Data_Table_object.Rows(Index__obj).Item("STATION") = Round(ChainageGrid0, nrdec)
                                                        Data_Table_object.Rows(Index__obj).Item("OFFSET") = Round(Offset1, nrdec)
                                                        Data_Table_object.Rows(Index__obj).Item("BLOCK_NAME") = "LINE_START"
                                                        Data_Table_object.Rows(Index__obj).Item("SEGMENT_NUMBER") = Segment_No


                                                        Dim Left_right As String = Angle_left_right(Poly2D, New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0), 1)
                                                        Data_Table_object.Rows(Index__obj).Item("LEFT_RIGHT") = Left_right



                                                        Dim Valoare As String = Pline_OBJ.Layer
                                                        Dim Valoare1 As String = ""
                                                        If IsNothing(Data_table_Layers) = False Then
                                                            If Data_table_Layers.Rows.Count > 0 Then
                                                                For i = 0 To Data_table_Layers.Rows.Count - 1
                                                                    If IsDBNull(Data_table_Layers.Rows(i).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(i).Item("DESCRIPTION")) = False Then
                                                                        If Data_table_Layers.Rows(i).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                            Valoare1 = Data_table_Layers.Rows(i).Item("DESCRIPTION")
                                                                            Exit For
                                                                        End If

                                                                    End If
                                                                Next
                                                            End If
                                                        End If

                                                        Data_Table_object.Rows(Index__obj).Item("LAYER_NAME") = Pline_OBJ.Layer
                                                        If Not Valoare1 = "" Then
                                                            Data_Table_object.Rows(Index__obj).Item("LAYER_DESCRIPTION") = Valoare1
                                                        End If


                                                        Index__obj = Index__obj + 1



                                                        Dim Station As String
                                                        If CheckBox_US_station.Checked = False Then
                                                            Station = Get_chainage_from_double(ChainageGrid0, 3)
                                                        Else
                                                            Station = Get_chainage_feet_from_double(ChainageGrid0, 0)
                                                        End If







                                                    End If



                                                Next

                                                Data_Table_object = Sort_data_table(Data_Table_object, "STATION")

                                                If Data_Table_object.Rows.Count > 0 Then

                                                    ReDim Preserve Line_new(Index_l - 1)
                                                    Line_new(Index_l - 1) = New Line
                                                    Line_new(Index_l - 1).StartPoint = New Point3d(Data_Table_object.Rows(0).Item("X1"), Data_Table_object.Rows(0).Item("Y1"), 0)
                                                    Line_new(Index_l - 1).EndPoint = New Point3d(Data_Table_object.Rows(0).Item("X"), Data_Table_object.Rows(0).Item("Y"), 0)
                                                    Line_new(Index_l - 1).Layer = "NO PLOT"

                                                    Index_l = Index_l + 1


                                                    Data_Table_information.Rows.Add()


                                                    Data_Table_information.Rows(Index_row).Item("X") = Data_Table_object.Rows(0).Item("X")
                                                    Data_Table_information.Rows(Index_row).Item("Y") = Data_Table_object.Rows(0).Item("Y")
                                                    Data_Table_information.Rows(Index_row).Item("Z") = 0
                                                    Data_Table_information.Rows(Index_row).Item("STATION") = Data_Table_object.Rows(0).Item("STATION")
                                                    Data_Table_information.Rows(Index_row).Item("OFFSET") = Data_Table_object.Rows(0).Item("OFFSET")
                                                    Data_Table_information.Rows(Index_row).Item("BLOCK_NAME") = "START"
                                                    Data_Table_information.Rows(Index_row).Item("SEGMENT_NUMBER") = Data_Table_object.Rows(0).Item("SEGMENT_NUMBER")



                                                    Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT") = Data_Table_object.Rows(0).Item("LEFT_RIGHT")




                                                    Data_Table_information.Rows(Index_row).Item("LAYER_NAME") = Data_Table_object.Rows(0).Item("LAYER_NAME")
                                                    If IsDBNull(Data_Table_object.Rows(0).Item("LAYER_DESCRIPTION")) = False Then
                                                        Data_Table_information.Rows(Index_row).Item("LAYER_DESCRIPTION") = Data_Table_object.Rows(0).Item("LAYER_DESCRIPTION")
                                                    End If


                                                    Dim mlED As New MLeader

                                                    mlED = Creaza_Mleader_nou_fara_UCS_transform_CU_btrecord_AND_TRANS(BTrecord, Trans1, New Point3d(Data_Table_information.Rows(Index_row).Item("X"), Data_Table_information.Rows(Index_row).Item("Y"), 0),
                                                                                                           "East = " & Round(Data_Table_information.Rows(Index_row).Item("X"), 3) & vbCrLf &
                                                                                                           "North = " & Round(Data_Table_information.Rows(Index_row).Item("Y"), 3) & vbCrLf &
                                                                                                           "Station = " & Data_Table_information.Rows(Index_row).Item("STATION") & vbCrLf &
                                                                                                           "Offset = " & Round(Data_Table_information.Rows(Index_row).Item("OFFSET"), 3) & Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT"), 1, 0.2, 0.5, 5, 5)



                                                    mlED.Layer = "NO PLOT"


                                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Object_ID_OD, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                        If IsNothing(Records1) = False Then
                                                            If Records1.Count > 0 Then
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

                                                                        If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                            Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                        End If
                                                                        If Not Replace(Valoare_field, " ", "") = "" Then
                                                                            Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If
                                                        End If
                                                    End Using



                                                    Index_row = Index_row + 1


                                                End If

                                                If Data_Table_object.Rows.Count > 1 Then
                                                    ReDim Preserve Line_new(Index_l - 1)
                                                    Line_new(Index_l - 1) = New Line
                                                    Line_new(Index_l - 1).StartPoint = New Point3d(Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("X1"), Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("Y1"), 0)
                                                    Line_new(Index_l - 1).EndPoint = New Point3d(Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("X"), Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("Y"), 0)
                                                    Line_new(Index_l - 1).Layer = "NO PLOT"

                                                    Index_l = Index_l + 1


                                                    Data_Table_information.Rows.Add()


                                                    Data_Table_information.Rows(Index_row).Item("X") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("X")
                                                    Data_Table_information.Rows(Index_row).Item("Y") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("Y")
                                                    Data_Table_information.Rows(Index_row).Item("Z") = 0
                                                    Data_Table_information.Rows(Index_row).Item("STATION") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("STATION")
                                                    Data_Table_information.Rows(Index_row).Item("OFFSET") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("OFFSET")
                                                    Data_Table_information.Rows(Index_row).Item("BLOCK_NAME") = "END"
                                                    Data_Table_information.Rows(Index_row).Item("SEGMENT_NUMBER") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("SEGMENT_NUMBER")



                                                    Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("LEFT_RIGHT")




                                                    Data_Table_information.Rows(Index_row).Item("LAYER_NAME") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("LAYER_NAME")
                                                    If IsDBNull(Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("LAYER_DESCRIPTION")) = False Then
                                                        Data_Table_information.Rows(Index_row).Item("LAYER_DESCRIPTION") = Data_Table_object.Rows(Data_Table_object.Rows.Count - 1).Item("LAYER_DESCRIPTION")
                                                    End If



                                                    Dim mlED As New MLeader

                                                    mlED = Creaza_Mleader_nou_fara_UCS_transform_CU_btrecord_AND_TRANS(BTrecord, Trans1, New Point3d(Data_Table_information.Rows(Index_row).Item("X"), Data_Table_information.Rows(Index_row).Item("Y"), 0),
                                                                                                           "East = " & Round(Data_Table_information.Rows(Index_row).Item("X"), 3) & vbCrLf &
                                                                                                           "North = " & Round(Data_Table_information.Rows(Index_row).Item("Y"), 3) & vbCrLf &
                                                                                                           "Station = " & Data_Table_information.Rows(Index_row).Item("STATION") & vbCrLf &
                                                                                                           "Offset = " & Round(Data_Table_information.Rows(Index_row).Item("OFFSET"), 3) & Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT"), 1, 0.2, 0.5, 5, 5)



                                                    mlED.Layer = "NO PLOT"

                                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Object_ID_OD, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                        If IsNothing(Records1) = False Then
                                                            If Records1.Count > 0 Then
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

                                                                        If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                            Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                        End If
                                                                        If Not Replace(Valoare_field, " ", "") = "" Then
                                                                            Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If
                                                        End If
                                                    End Using


                                                    Index_row = Index_row + 1
                                                End If

                                                Segment_No = Segment_No + 1





                                            End If
                                        End If
                                    End If


                                Next


                                If IsNothing(Line_new) = False Then
                                    If Line_new.Length > 0 Then
                                        For Jj = 0 To Line_new.Length - 1

                                            BTrecord.AppendEntity(Line_new(Jj))
                                            Trans1.AddNewlyCreatedDBObject(Line_new(Jj), True)
                                        Next
                                    End If
                                End If




                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                If Data_Table_information.Rows.Count > 0 Then
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                    For i = 0 To Data_Table_information.Columns.Count - 1
                        W1.Cells(1, i + 1).value2 = Data_Table_information.Columns(i).ColumnName
                    Next
                    Dim Rand_excel As Double = 2
                    For i = 0 To Data_Table_information.Rows.Count - 1
                        For j = 0 To Data_Table_information.Columns.Count - 1
                            If IsDBNull(Data_Table_information.Rows(i).Item(j)) = False Then
                                W1.Cells(Rand_excel, j + 1).value2 = Data_Table_information.Rows(i).Item(j)
                            End If
                        Next
                        Rand_excel = Rand_excel + 1
                    Next
                End If

                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Private Sub Button_residence_scanning_Click(sender As Object, e As EventArgs) Handles Button_residence_scanning.Click

        If Freeze_operations = False Then

            Try






                Dim Data_Table_information As New System.Data.DataTable
                Data_Table_information.Columns.Add("STATION", GetType(Double))
                Data_Table_information.Columns.Add("DESCRIPTION", GetType(String))
                Data_Table_information.Columns.Add("OFFSET_CL", GetType(Double))
                Data_Table_information.Columns.Add("OFFSET_LOC", GetType(Double))
                Data_Table_information.Columns.Add("LEFT_RIGHT", GetType(String))
                Data_Table_information.Columns.Add("LAYER_NAME", GetType(String))
                Data_Table_information.Columns.Add("X_CL", GetType(Double))
                Data_Table_information.Columns.Add("Y_CL", GetType(Double))
                Data_Table_information.Columns.Add("X_LOC", GetType(Double))
                Data_Table_information.Columns.Add("Y_LOC", GetType(Double))
                Data_Table_information.Columns.Add("X0", GetType(Double))
                Data_Table_information.Columns.Add("Y0", GetType(Double))
                Data_Table_information.Columns.Add("NOTES", GetType(String))

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select centerline:"

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select limits of construction polyline:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)


                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False And IsNothing(Rezultat2) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                Dim Poly3d_CL As Polyline3d = Nothing
                                Dim Poly3d_LOC As Polyline3d = Nothing

                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)
                                Dim Ent2 As Entity
                                Ent2 = Trans1.GetObject(Rezultat2.Value.Item(0).ObjectId, OpenMode.ForRead)

                                Dim Index_row As Double = 0


                                Dim Poly2D_CL As New Polyline
                                Dim Poly2D_LOC As New Polyline

                                If TypeOf Ent1 Is Polyline3d Then
                                    Poly3d_CL = Ent1
                                    Dim Index2d As Double = 0
                                    For Each ObjId As ObjectId In Poly3d_CL
                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                        Poly2D_CL.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                        Index2d = Index2d + 1
                                    Next
                                End If

                                If TypeOf Ent2 Is Polyline3d Then
                                    Poly3d_LOC = Ent2
                                    Dim Index2d As Double = 0
                                    For Each ObjId As ObjectId In Poly3d_LOC
                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                        Poly2D_LOC.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                        Index2d = Index2d + 1
                                    Next
                                End If

                                If TypeOf Ent1 Is Polyline Then
                                    Poly2D_CL = Ent1.Clone
                                End If

                                If TypeOf Ent2 Is Polyline Then
                                    Poly2D_LOC = Ent2.Clone
                                End If


                                If Poly2D_CL = Nothing Or Poly2D_LOC = Nothing Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Poly2D_CL.Elevation = 0
                                Poly2D_LOC.Elevation = 0


                                Dim Data_table_Columns_object_data As New System.Data.DataTable
                                Data_table_Columns_object_data.Columns.Add("NAME", GetType(String))
                                Data_table_Columns_object_data.Columns.Add("NR", GetType(Integer))
                                Dim nrdec As Integer = 3
                                If CheckBox_zero_decimals.Checked = True Then nrdec = 0


                                Dim Line_new_CL() As Autodesk.AutoCAD.DatabaseServices.Line
                                Dim Index_l As Integer = 1

                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables


                                For Each ObjID In BTrecord

                                    Dim Ent_Object As Entity = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                                    If Not Ent_Object = Nothing Then
                                        If Not ObjID = Ent1.ObjectId And Not ObjID = Ent2.ObjectId Then

                                            If TypeOf (Ent_Object) Is Polyline Or TypeOf (Ent_Object) Is Line Then
                                                Dim Object_ID_OD As ObjectId

                                                Dim Line_OBJ As Autodesk.AutoCAD.DatabaseServices.Line = TryCast(Ent_Object, Autodesk.AutoCAD.DatabaseServices.Line)

                                                Dim Pline_OBJ As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(Ent_Object, Autodesk.AutoCAD.DatabaseServices.Polyline)

                                                If IsNothing(Line_OBJ) = False Then
                                                    Pline_OBJ = New Polyline
                                                    Pline_OBJ.AddVertexAt(0, New Point2d(Line_OBJ.StartPoint.X, Line_OBJ.StartPoint.Y), 0, 0, 0)
                                                    Pline_OBJ.AddVertexAt(1, New Point2d(Line_OBJ.EndPoint.X, Line_OBJ.EndPoint.Y), 0, 0, 0)
                                                    Object_ID_OD = Line_OBJ.ObjectId
                                                Else
                                                    Object_ID_OD = Pline_OBJ.ObjectId
                                                End If


                                                Dim Data_Table_object As New System.Data.DataTable
                                                Data_Table_object.Columns.Add("STATION", GetType(Double))
                                                Data_Table_object.Columns.Add("DESCRIPTION", GetType(String))

                                                Data_Table_object.Columns.Add("OFFSET_LOC", GetType(Double))
                                                Data_Table_object.Columns.Add("OFFSET_CL", GetType(Double))

                                                Data_Table_object.Columns.Add("X_CL", GetType(Double))
                                                Data_Table_object.Columns.Add("Y_CL", GetType(Double))


                                                Data_Table_object.Columns.Add("LAYER_NAME", GetType(String))

                                                Data_Table_object.Columns.Add("LEFT_RIGHT", GetType(String))
                                                Data_Table_object.Columns.Add("X_LOC", GetType(Double))
                                                Data_Table_object.Columns.Add("Y_LOC", GetType(Double))


                                                Data_Table_object.Columns.Add("X0", GetType(Double))
                                                Data_Table_object.Columns.Add("Y0", GetType(Double))


                                                Dim Index__obj As Double = 0
                                                For j = 0 To Pline_OBJ.NumberOfVertices - 1




                                                    Dim Point_on_poly2d_CL As New Point3d
                                                    Dim Point_on_poly2d_LOC As New Point3d




                                                    Point_on_poly2d_CL = Poly2D_CL.GetClosestPointTo(New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0), Vector3d.ZAxis, True)

                                                    Dim Offset_CL As Double = New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0).GetVectorTo(Point_on_poly2d_CL).Length

                                                    Point_on_poly2d_LOC = Poly2D_LOC.GetClosestPointTo(New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0), Vector3d.ZAxis, True)

                                                    Dim Offset_LOC As Double = New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0).GetVectorTo(Point_on_poly2d_LOC).Length



                                                    'If Offset_LOC <= CDbl(TextBox_buffer.Text) Then

                                                    Dim ChainageGrid0 As Double

                                                    If Not Poly3d_CL = Nothing Then

                                                        Dim Point_on_poly3d As New Point3d
                                                        Dim Param2d As Double = Poly2D_CL.GetParameterAtPoint(Point_on_poly2d_CL)
                                                        Point_on_poly3d = Poly3d_CL.GetPointAtParameter(Param2d)

                                                        ChainageGrid0 = Poly3d_CL.GetDistAtPoint(Point_on_poly3d)
                                                    Else

                                                        ChainageGrid0 = Poly2D_CL.GetDistAtPoint(Point_on_poly2d_CL)
                                                    End If




                                                    Data_Table_object.Rows.Add()
                                                    Data_Table_object.Rows(Index__obj).Item("X0") = Round(Pline_OBJ.GetPoint3dAt(j).X, 3)
                                                    Data_Table_object.Rows(Index__obj).Item("Y0") = Round(Pline_OBJ.GetPoint3dAt(j).Y, 3)


                                                    Data_Table_object.Rows(Index__obj).Item("X_CL") = Round(Point_on_poly2d_CL.X, 3)
                                                    Data_Table_object.Rows(Index__obj).Item("Y_CL") = Round(Point_on_poly2d_CL.Y, 3)

                                                    Data_Table_object.Rows(Index__obj).Item("STATION") = Round(ChainageGrid0, nrdec)
                                                    Data_Table_object.Rows(Index__obj).Item("OFFSET_CL") = Round(Offset_CL, nrdec)
                                                    Data_Table_object.Rows(Index__obj).Item("OFFSET_LOC") = Round(Offset_LOC, nrdec)

                                                    Data_Table_object.Rows(Index__obj).Item("X_LOC") = Round(Point_on_poly2d_LOC.X, 3)
                                                    Data_Table_object.Rows(Index__obj).Item("Y_LOC") = Round(Point_on_poly2d_LOC.Y, 3)



                                                    Dim Left_right As String = Angle_left_right(Poly2D_CL, New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0), 1)
                                                    Data_Table_object.Rows(Index__obj).Item("LEFT_RIGHT") = Left_right




                                                    Data_Table_object.Rows(Index__obj).Item("LAYER_NAME") = Pline_OBJ.Layer



                                                    Index__obj = Index__obj + 1



                                                    Dim Station As String
                                                    If CheckBox_US_station.Checked = False Then
                                                        Station = Get_chainage_from_double(ChainageGrid0, 3)
                                                    Else
                                                        Station = Get_chainage_feet_from_double(ChainageGrid0, 0)
                                                    End If











                                                Next

                                                Dim Data_Table_object_loc As New System.Data.DataTable

                                                Data_Table_object_loc = Sort_data_table(Data_Table_object, "OFFSET_LOC")


                                                Dim Data_Table_object_cl As New System.Data.DataTable

                                                Data_Table_object_cl = Sort_data_table(Data_Table_object, "OFFSET_CL")



                                                If Data_Table_object.Rows.Count > 0 Then

                                                    ReDim Preserve Line_new_CL(Index_l - 1)
                                                    Line_new_CL(Index_l - 1) = New Line
                                                    Line_new_CL(Index_l - 1).StartPoint = New Point3d(Data_Table_object_cl.Rows(0).Item("X0"), Data_Table_object_cl.Rows(0).Item("Y0"), 0)
                                                    Line_new_CL(Index_l - 1).EndPoint = New Point3d(Data_Table_object_cl.Rows(0).Item("X_CL"), Data_Table_object_cl.Rows(0).Item("Y_CL"), 0)
                                                    Line_new_CL(Index_l - 1).Layer = "NO PLOT"









                                                    Data_Table_information.Rows.Add()


                                                    Data_Table_information.Rows(Index_row).Item("X_CL") = Data_Table_object_cl.Rows(0).Item("X_CL")
                                                    Data_Table_information.Rows(Index_row).Item("Y_CL") = Data_Table_object_cl.Rows(0).Item("Y_CL")

                                                    Data_Table_information.Rows(Index_row).Item("STATION") = Data_Table_object_cl.Rows(0).Item("STATION")

                                                    Data_Table_information.Rows(Index_row).Item("X_LOC") = Data_Table_object_loc.Rows(0).Item("X_LOC")
                                                    Data_Table_information.Rows(Index_row).Item("Y_LOC") = Data_Table_object_loc.Rows(0).Item("Y_LOC")


                                                    Data_Table_information.Rows(Index_row).Item("X0") = Data_Table_object_cl.Rows(0).Item("X0")
                                                    Data_Table_information.Rows(Index_row).Item("Y0") = Data_Table_object_cl.Rows(0).Item("Y0")


                                                    Data_Table_information.Rows(Index_row).Item("OFFSET_CL") = Data_Table_object_cl.Rows(0).Item("OFFSET_CL")


                                                    Data_Table_information.Rows(Index_row).Item("OFFSET_LOC") = Data_Table_object_loc.Rows(0).Item("OFFSET_LOC")




                                                    If Not Data_Table_object_cl.Rows(0).Item("X0") = Data_Table_object_loc.Rows(0).Item("X0") Or
                                                        Not Data_Table_object_cl.Rows(0).Item("Y0") = Data_Table_object_loc.Rows(0).Item("Y0") Or
                                                        Not Round(GET_Bearing_rad(Data_Table_object_cl.Rows(0).Item("X0"), Data_Table_object_cl.Rows(0).Item("Y0"), Data_Table_object_cl.Rows(0).Item("X_CL"), Data_Table_object_cl.Rows(0).Item("Y_CL")), 2) =
                                                        Round(GET_Bearing_rad(Data_Table_object_cl.Rows(0).Item("X0"), Data_Table_object_cl.Rows(0).Item("Y0"), Data_Table_object_cl.Rows(0).Item("X_LOC"), Data_Table_object_cl.Rows(0).Item("Y_LOC")), 2) Then

                                                        Data_Table_information.Rows(Index_row).Item("NOTES") = "CHECK"



                                                        Index_l = Index_l + 1

                                                        ReDim Preserve Line_new_CL(Index_l - 1)

                                                        Line_new_CL(Index_l - 1) = New Line
                                                        Line_new_CL(Index_l - 1).StartPoint = New Point3d(Data_Table_object_loc.Rows(0).Item("X0"), Data_Table_object_loc.Rows(0).Item("Y0"), 0)
                                                        Line_new_CL(Index_l - 1).EndPoint = New Point3d(Data_Table_object_loc.Rows(0).Item("X_LOC"), Data_Table_object_loc.Rows(0).Item("Y_LOC"), 0)
                                                        Line_new_CL(Index_l - 1).Layer = "NO PLOT"


                                                    End If
                                                    Index_l = Index_l + 1




                                                    Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT") = Data_Table_object_cl.Rows(0).Item("LEFT_RIGHT")




                                                    Data_Table_information.Rows(Index_row).Item("LAYER_NAME") = Data_Table_object_loc.Rows(0).Item("LAYER_NAME")



                                                    Dim mlED As New MLeader

                                                    mlED = Creaza_Mleader_nou_fara_UCS_transform_CU_btrecord_AND_TRANS(BTrecord, Trans1, New Point3d(Data_Table_information.Rows(Index_row).Item("X_CL"), Data_Table_information.Rows(Index_row).Item("Y_CL"), 0),
                                                                                                           "Station = " & Data_Table_information.Rows(Index_row).Item("STATION") & vbCrLf &
                                                                                                           "Offset = " & Round(Data_Table_information.Rows(Index_row).Item("OFFSET_CL"), 3) & Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT"), 1, 0.2, 0.5, 5, 5)



                                                    mlED.Layer = "NO PLOT"


                                                    Dim mlED1 As New MLeader

                                                    mlED1 = Creaza_Mleader_nou_fara_UCS_transform_CU_btrecord_AND_TRANS(BTrecord, Trans1, New Point3d(Data_Table_information.Rows(Index_row).Item("X_LOC"), Data_Table_information.Rows(Index_row).Item("Y_LOC"), 0),
                                                                                                           "Offset = " & Round(Data_Table_information.Rows(Index_row).Item("OFFSET_LOC"), 3) & Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT"), 1, 0.2, 0.5, 5, 10)



                                                    mlED1.Layer = "NO PLOT"



                                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Object_ID_OD, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                        If IsNothing(Records1) = False Then
                                                            If Records1.Count > 0 Then
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

                                                                        If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                            Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                        End If
                                                                        If Not Replace(Valoare_field, " ", "") = "" Then
                                                                            Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If
                                                        End If
                                                    End Using



                                                    Index_row = Index_row + 1


                                                End If









                                            End If
                                        End If
                                    End If


                                Next


                                If IsNothing(Line_new_CL) = False Then
                                    If Line_new_CL.Length > 0 Then
                                        For Jj = 0 To Line_new_CL.Length - 1

                                            BTrecord.AppendEntity(Line_new_CL(Jj))
                                            Trans1.AddNewlyCreatedDBObject(Line_new_CL(Jj), True)
                                        Next
                                    End If
                                End If




                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                If Data_Table_information.Rows.Count > 0 Then
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                    For i = 0 To Data_Table_information.Columns.Count - 1
                        W1.Cells(1, i + 1).value2 = Data_Table_information.Columns(i).ColumnName
                    Next
                    Dim Rand_excel As Double = 2
                    For i = 0 To Data_Table_information.Rows.Count - 1
                        For j = 0 To Data_Table_information.Columns.Count - 1
                            If IsDBNull(Data_Table_information.Rows(i).Item(j)) = False Then
                                W1.Cells(Rand_excel, j + 1).value2 = Data_Table_information.Rows(i).Item(j)
                            End If
                        Next
                        Rand_excel = Rand_excel + 1
                    Next
                End If

                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Private Sub Button_poly_3d_face_scanning_Click(sender As Object, e As EventArgs) Handles Button_poly_3d_face_scanning.Click
        If Freeze_operations = False Then

            Try




                Dim Data_Table_information As New System.Data.DataTable
                Data_Table_information.Columns.Add("NUMBER", GetType(Double))
                Data_Table_information.Columns.Add("X", GetType(Double))
                Data_Table_information.Columns.Add("Y", GetType(Double))


                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select Polyline and Mtext objects"

                    Object_Prompt.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)



                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)



                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Index1 As Integer = 0

                                For i = 0 To Rezultat1.Value.Count - 1
                                    Dim Ent1 As Entity
                                    Ent1 = Trans1.GetObject(Rezultat1.Value.Item(i).ObjectId, OpenMode.ForRead)
                                    If TypeOf Ent1 Is MText Then
                                        Dim Mtext1 As MText = Ent1
                                        If IsNumeric(Mtext1.Text) = True Then
                                            Dim Dist1 As Double = 1000
                                            Dim ObjId1 As ObjectId = ObjectId.Null

                                            For Each OBJID As ObjectId In BTrecord
                                                If Not OBJID = Ent1.ObjectId Then
                                                    Dim Ent2 As Entity = Trans1.GetObject(OBJID, OpenMode.ForRead)
                                                    If TypeOf Ent2 Is Polyline Then
                                                        Dim Poly1 As New Polyline
                                                        Poly1 = Ent2.Clone
                                                        Poly1.Elevation = Mtext1.Location.Z
                                                        Dim Len1 As Double = Mtext1.Location.GetVectorTo(Poly1.GetClosestPointTo(Mtext1.Location, Vector3d.ZAxis, False)).Length
                                                        If Len1 < Dist1 Then
                                                            Dist1 = Len1
                                                            ObjId1 = OBJID
                                                        End If
                                                    End If
                                                End If
                                            Next

                                            If Not ObjId1 = ObjectId.Null Then
                                                Dim Poly1 As Polyline
                                                Poly1 = Trans1.GetObject(ObjId1, OpenMode.ForRead)
                                                Dim POint1 As New Point3d
                                                Data_Table_information.Rows.Add()
                                                For j = 0 To Poly1.NumberOfVertices - 1
                                                    If j = 0 Then
                                                        Data_Table_information.Rows(Index1).Item("X") = Poly1.GetPointAtParameter(j).X
                                                        Data_Table_information.Rows(Index1).Item("Y") = Poly1.GetPointAtParameter(j).Y
                                                        Data_Table_information.Rows(Index1).Item("NUMBER") = Mtext1.Text

                                                    Else
                                                        Dim X As Double = Data_Table_information.Rows(Index1).Item("X")
                                                        Dim Y As Double = Data_Table_information.Rows(Index1).Item("Y")
                                                        If Poly1.GetPointAtParameter(j).Y > Y Then
                                                            Data_Table_information.Rows(Index1).Item("Y") = Poly1.GetPointAtParameter(j).Y
                                                            Data_Table_information.Rows(Index1).Item("X") = Poly1.GetPointAtParameter(j).X

                                                        End If

                                                    End If


                                                Next
                                                Index1 = Index1 + 1
                                            End If

                                        End If
                                    End If
                                Next






                                If Data_Table_information.Rows.Count > 0 Then
                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                                    For i = 0 To Data_Table_information.Columns.Count - 1
                                        W1.Cells(1, i + 1).value2 = Data_Table_information.Columns(i).ColumnName
                                    Next
                                    Dim Rand_excel As Double = 2
                                    For i = 0 To Data_Table_information.Rows.Count - 1
                                        For j = 0 To Data_Table_information.Columns.Count - 1
                                            If IsDBNull(Data_Table_information.Rows(i).Item(j)) = False Then
                                                W1.Cells(Rand_excel, j + 1).value2 = Data_Table_information.Rows(i).Item(j)
                                            End If
                                        Next
                                        If IsDBNull(Data_Table_information.Rows(i).Item("X")) = False And IsDBNull(Data_Table_information.Rows(i).Item("Y")) = False And IsDBNull(Data_Table_information.Rows(i).Item("NUMBER")) = False Then
                                            Dim pT1 As New Point3d(Data_Table_information.Rows(i).Item("X"), Data_Table_information.Rows(i).Item("Y"), 0)
                                            Dim mLEADER1 As New MLeader

                                            mLEADER1 = Creaza_Mleader_nou_fara_UCS_transform(pT1, Data_Table_information.Rows(i).Item("NUMBER"), 1, 0.5, 0.25, 2, 5)
                                            mLEADER1.Layer = "NO PLOT"

                                        End If


                                        Rand_excel = Rand_excel + 1
                                    Next
                                End If











                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using



                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Private Sub Button_scan_single_point_segment_Click(sender As Object, e As EventArgs) Handles Button_scan_single_point_segment.Click
        If Freeze_operations = False Then

            Try



                If IsNumeric(TextBox_buffer.Text) = False Then
                    MsgBox("Please specify the Buffer size!")
                    Exit Sub
                End If


                Dim Data_Table_information As New System.Data.DataTable
                Data_Table_information.Columns.Add("STATION", GetType(Double))
                Data_Table_information.Columns.Add("DESCRIPTION", GetType(String))



                Data_Table_information.Columns.Add("X", GetType(Double))
                Data_Table_information.Columns.Add("Y", GetType(Double))
                Data_Table_information.Columns.Add("Z", GetType(Double))


                Data_Table_information.Columns.Add("LAYER_NAME", GetType(String))
                Data_Table_information.Columns.Add("LAYER_DESCRIPTION", GetType(String))


                Data_Table_information.Columns.Add("OFFSET", GetType(Double))

                Data_Table_information.Columns.Add("LEFT_RIGHT", GetType(String))



                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select centerline:"

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                Dim Poly3d As Polyline3d = Nothing
                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)





                                Dim Index_row As Double = 0

                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection


                                Dim Poly2D As New Polyline


                                If TypeOf Ent1 Is Polyline3d Then
                                    Poly3d = Ent1
                                    Dim Index2d As Double = 0
                                    For Each ObjId As ObjectId In Poly3d
                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                        Index2d = Index2d + 1
                                    Next
                                End If


                                If TypeOf Ent1 Is Polyline Then
                                    Poly2D = Ent1.Clone
                                End If

                                If Poly2D = Nothing Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Poly2D.Elevation = 0




                                Dim Data_table_Columns_object_data As New System.Data.DataTable
                                Data_table_Columns_object_data.Columns.Add("NAME", GetType(String))
                                Data_table_Columns_object_data.Columns.Add("NR", GetType(Integer))
                                Dim nrdec As Integer = 3
                                If CheckBox_zero_decimals.Checked = True Then nrdec = 0


                                Dim Line_new() As Autodesk.AutoCAD.DatabaseServices.Line
                                Dim Index_l As Integer = 1

                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables


                                For Each ObjID In BTrecord

                                    Dim Ent_Object As Entity = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                                    If Not Ent_Object = Nothing Then
                                        If Not ObjID = Ent1.ObjectId Then

                                            If TypeOf (Ent_Object) Is Polyline Or TypeOf (Ent_Object) Is Line Then
                                                Dim Object_ID_OD As ObjectId

                                                Dim Line_OBJ As Autodesk.AutoCAD.DatabaseServices.Line = TryCast(Ent_Object, Autodesk.AutoCAD.DatabaseServices.Line)

                                                Dim Pline_OBJ As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(Ent_Object, Autodesk.AutoCAD.DatabaseServices.Polyline)

                                                If IsNothing(Line_OBJ) = False Then
                                                    Pline_OBJ = New Polyline
                                                    Pline_OBJ.AddVertexAt(0, New Point2d(Line_OBJ.StartPoint.X, Line_OBJ.StartPoint.Y), 0, 0, 0)
                                                    Pline_OBJ.AddVertexAt(1, New Point2d(Line_OBJ.EndPoint.X, Line_OBJ.EndPoint.Y), 0, 0, 0)
                                                    Object_ID_OD = Line_OBJ.ObjectId
                                                Else
                                                    Object_ID_OD = Pline_OBJ.ObjectId
                                                End If


                                                Dim Data_Table_object As New System.Data.DataTable
                                                Data_Table_object.Columns.Add("X", GetType(Double))
                                                Data_Table_object.Columns.Add("Y", GetType(Double))
                                                Data_Table_object.Columns.Add("Z", GetType(Double))

                                                Data_Table_object.Columns.Add("STATION", GetType(Double))

                                                Data_Table_object.Columns.Add("LAYER_NAME", GetType(String))
                                                Data_Table_object.Columns.Add("LAYER_DESCRIPTION", GetType(String))


                                                Data_Table_object.Columns.Add("OFFSET", GetType(Double))

                                                Data_Table_object.Columns.Add("LEFT_RIGHT", GetType(String))
                                                Data_Table_object.Columns.Add("X1", GetType(Double))
                                                Data_Table_object.Columns.Add("Y1", GetType(Double))
                                                Data_Table_object.Columns.Add("Z1", GetType(Double))
                                                Dim Index__obj As Double = 0
                                                For j = 0 To Pline_OBJ.NumberOfVertices - 1


                                                    Dim Point_on_poly As New Point3d
                                                    Dim Point_on_poly3d As New Point3d
                                                    Dim Point_on_poly2d As New Point3d


                                                    Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0), Vector3d.ZAxis, True)

                                                    Dim Offset1 As Double = New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0).GetVectorTo(Point_on_poly2d).Length





                                                    If Offset1 <= CDbl(TextBox_buffer.Text) Then

                                                        Dim ChainageGrid0 As Double

                                                        If Not Poly3d = Nothing Then
                                                            Dim Param2d As Double = Poly2D.GetParameterAtPoint(Point_on_poly2d)
                                                            Point_on_poly3d = Poly3d.GetPointAtParameter(Param2d)
                                                            Point_on_poly = Point_on_poly3d
                                                            ChainageGrid0 = Poly3d.GetDistAtPoint(Point_on_poly3d)
                                                        Else
                                                            Point_on_poly = Point_on_poly2d
                                                            ChainageGrid0 = Poly2D.GetDistAtPoint(Point_on_poly2d)
                                                        End If




                                                        Data_Table_object.Rows.Add()
                                                        Data_Table_object.Rows(Index__obj).Item("X1") = Round(Pline_OBJ.GetPoint3dAt(j).X, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Y1") = Round(Pline_OBJ.GetPoint3dAt(j).Y, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Z1") = 0

                                                        Data_Table_object.Rows(Index__obj).Item("X") = Round(Point_on_poly.X, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Y") = Round(Point_on_poly.Y, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Z") = 0
                                                        Data_Table_object.Rows(Index__obj).Item("STATION") = Round(ChainageGrid0, nrdec)
                                                        Data_Table_object.Rows(Index__obj).Item("OFFSET") = Round(Offset1, nrdec)




                                                        Dim Left_right As String = Angle_left_right(Poly2D, New Point3d(Pline_OBJ.GetPoint3dAt(j).X, Pline_OBJ.GetPoint3dAt(j).Y, 0), 1)
                                                        Data_Table_object.Rows(Index__obj).Item("LEFT_RIGHT") = Left_right



                                                        Dim Valoare As String = Pline_OBJ.Layer
                                                        Dim Valoare1 As String = ""
                                                        If IsNothing(Data_table_Layers) = False Then
                                                            If Data_table_Layers.Rows.Count > 0 Then
                                                                For i = 0 To Data_table_Layers.Rows.Count - 1
                                                                    If IsDBNull(Data_table_Layers.Rows(i).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(i).Item("DESCRIPTION")) = False Then
                                                                        If Data_table_Layers.Rows(i).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                            Valoare1 = Data_table_Layers.Rows(i).Item("DESCRIPTION")
                                                                            Exit For
                                                                        End If

                                                                    End If
                                                                Next
                                                            End If
                                                        End If

                                                        Data_Table_object.Rows(Index__obj).Item("LAYER_NAME") = Pline_OBJ.Layer
                                                        If Not Valoare1 = "" Then
                                                            Data_Table_object.Rows(Index__obj).Item("LAYER_DESCRIPTION") = Valoare1
                                                        End If


                                                        Index__obj = Index__obj + 1



                                                        Dim Station As String
                                                        If CheckBox_US_station.Checked = False Then
                                                            Station = Get_chainage_from_double(ChainageGrid0, 3)
                                                        Else
                                                            Station = Get_chainage_feet_from_double(ChainageGrid0, 0)
                                                        End If







                                                    End If



                                                Next

                                                Data_Table_object = Sort_data_table(Data_Table_object, "OFFSET")

                                                If Data_Table_object.Rows.Count > 0 Then

                                                    ReDim Preserve Line_new(Index_l - 1)
                                                    Line_new(Index_l - 1) = New Line
                                                    Line_new(Index_l - 1).StartPoint = New Point3d(Data_Table_object.Rows(0).Item("X1"), Data_Table_object.Rows(0).Item("Y1"), 0)
                                                    Line_new(Index_l - 1).EndPoint = New Point3d(Data_Table_object.Rows(0).Item("X"), Data_Table_object.Rows(0).Item("Y"), 0)
                                                    Line_new(Index_l - 1).Layer = "NO PLOT"

                                                    Index_l = Index_l + 1


                                                    Data_Table_information.Rows.Add()


                                                    Data_Table_information.Rows(Index_row).Item("X") = Data_Table_object.Rows(0).Item("X")
                                                    Data_Table_information.Rows(Index_row).Item("Y") = Data_Table_object.Rows(0).Item("Y")
                                                    Data_Table_information.Rows(Index_row).Item("Z") = 0
                                                    Data_Table_information.Rows(Index_row).Item("STATION") = Data_Table_object.Rows(0).Item("STATION")
                                                    Data_Table_information.Rows(Index_row).Item("OFFSET") = Data_Table_object.Rows(0).Item("OFFSET")





                                                    Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT") = Data_Table_object.Rows(0).Item("LEFT_RIGHT")




                                                    Data_Table_information.Rows(Index_row).Item("LAYER_NAME") = Data_Table_object.Rows(0).Item("LAYER_NAME")
                                                    If IsDBNull(Data_Table_object.Rows(0).Item("LAYER_DESCRIPTION")) = False Then
                                                        Data_Table_information.Rows(Index_row).Item("LAYER_DESCRIPTION") = Data_Table_object.Rows(0).Item("LAYER_DESCRIPTION")
                                                    End If


                                                    Dim mlED As New MLeader

                                                    mlED = Creaza_Mleader_nou_fara_UCS_transform_CU_btrecord_AND_TRANS(BTrecord, Trans1, New Point3d(Data_Table_information.Rows(Index_row).Item("X"), Data_Table_information.Rows(Index_row).Item("Y"), 0),
                                                                                                           "East = " & Round(Data_Table_information.Rows(Index_row).Item("X"), 3) & vbCrLf &
                                                                                                           "North = " & Round(Data_Table_information.Rows(Index_row).Item("Y"), 3) & vbCrLf &
                                                                                                           "Station = " & Data_Table_information.Rows(Index_row).Item("STATION") & vbCrLf &
                                                                                                           "Offset = " & Round(Data_Table_information.Rows(Index_row).Item("OFFSET"), 3) & Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT"), 1, 0.2, 0.5, 5, 5)



                                                    mlED.Layer = "NO PLOT"


                                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Object_ID_OD, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                        If IsNothing(Records1) = False Then
                                                            If Records1.Count > 0 Then
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

                                                                        If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                            Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                        End If
                                                                        If Not Replace(Valoare_field, " ", "") = "" Then
                                                                            Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If
                                                        End If
                                                    End Using



                                                    Index_row = Index_row + 1


                                                End If



                                            End If

                                            If TypeOf (Ent_Object) Is Circle Then
                                                Dim Object_ID_OD As ObjectId

                                                Dim Circle_OBJ As Autodesk.AutoCAD.DatabaseServices.Circle = TryCast(Ent_Object, Autodesk.AutoCAD.DatabaseServices.Circle)

                                                If IsNothing(Circle_OBJ) = False Then

                                                    Object_ID_OD = Circle_OBJ.ObjectId



                                                    Dim Data_Table_object As New System.Data.DataTable
                                                    Data_Table_object.Columns.Add("X", GetType(Double))
                                                    Data_Table_object.Columns.Add("Y", GetType(Double))
                                                    Data_Table_object.Columns.Add("Z", GetType(Double))

                                                    Data_Table_object.Columns.Add("STATION", GetType(Double))

                                                    Data_Table_object.Columns.Add("LAYER_NAME", GetType(String))
                                                    Data_Table_object.Columns.Add("LAYER_DESCRIPTION", GetType(String))


                                                    Data_Table_object.Columns.Add("OFFSET", GetType(Double))

                                                    Data_Table_object.Columns.Add("LEFT_RIGHT", GetType(String))
                                                    Data_Table_object.Columns.Add("X1", GetType(Double))
                                                    Data_Table_object.Columns.Add("Y1", GetType(Double))
                                                    Data_Table_object.Columns.Add("Z1", GetType(Double))
                                                    Dim Index__obj As Double = 0






                                                    Dim Point_on_poly As New Point3d
                                                    Dim Point_on_poly3d As New Point3d
                                                    Dim Point_on_circle As New Point3d


                                                    Point_on_circle = Poly2D.GetClosestPointTo(New Point3d(Circle_OBJ.Center.X, Circle_OBJ.Center.Y, 0), Vector3d.ZAxis, True)

                                                    Dim Offset1 As Double = New Point3d(Circle_OBJ.Center.X, Circle_OBJ.Center.Y, 0).GetVectorTo(Point_on_circle).Length





                                                    If Offset1 <= CDbl(TextBox_buffer.Text) Then

                                                        Dim ChainageGrid0 As Double

                                                        If Not Poly3d = Nothing Then
                                                            Dim Param2d As Double = Poly2D.GetParameterAtPoint(Point_on_circle)
                                                            Point_on_poly3d = Poly3d.GetPointAtParameter(Param2d)
                                                            Point_on_poly = Point_on_poly3d
                                                            ChainageGrid0 = Poly3d.GetDistAtPoint(Point_on_poly3d)
                                                        Else
                                                            Point_on_poly = Point_on_circle
                                                            ChainageGrid0 = Poly2D.GetDistAtPoint(Point_on_circle)
                                                        End If




                                                        Data_Table_object.Rows.Add()
                                                        Data_Table_object.Rows(Index__obj).Item("X1") = Round(Circle_OBJ.Center.X, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Y1") = Round(Circle_OBJ.Center.Y, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Z1") = 0

                                                        Data_Table_object.Rows(Index__obj).Item("X") = Round(Point_on_poly.X, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Y") = Round(Point_on_poly.Y, 3)
                                                        Data_Table_object.Rows(Index__obj).Item("Z") = 0
                                                        Data_Table_object.Rows(Index__obj).Item("STATION") = Round(ChainageGrid0, nrdec)
                                                        Data_Table_object.Rows(Index__obj).Item("OFFSET") = Round(Offset1, nrdec)




                                                        Dim Left_right As String = Angle_left_right(Poly2D, New Point3d(Circle_OBJ.Center.X, Circle_OBJ.Center.Y, 0), 1)
                                                        Data_Table_object.Rows(Index__obj).Item("LEFT_RIGHT") = Left_right



                                                        Dim Valoare As String = Circle_OBJ.Layer
                                                        Dim Valoare1 As String = ""
                                                        If IsNothing(Data_table_Layers) = False Then
                                                            If Data_table_Layers.Rows.Count > 0 Then
                                                                For i = 0 To Data_table_Layers.Rows.Count - 1
                                                                    If IsDBNull(Data_table_Layers.Rows(i).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(i).Item("DESCRIPTION")) = False Then
                                                                        If Data_table_Layers.Rows(i).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                            Valoare1 = Data_table_Layers.Rows(i).Item("DESCRIPTION")
                                                                            Exit For
                                                                        End If

                                                                    End If
                                                                Next
                                                            End If
                                                        End If

                                                        Data_Table_object.Rows(Index__obj).Item("LAYER_NAME") = Circle_OBJ.Layer
                                                        If Not Valoare1 = "" Then
                                                            Data_Table_object.Rows(Index__obj).Item("LAYER_DESCRIPTION") = Valoare1
                                                        End If


                                                        Index__obj = Index__obj + 1



                                                        Dim Station As String
                                                        If CheckBox_US_station.Checked = False Then
                                                            Station = Get_chainage_from_double(ChainageGrid0, 3)
                                                        Else
                                                            Station = Get_chainage_feet_from_double(ChainageGrid0, 0)
                                                        End If







                                                    End If





                                                    Data_Table_object = Sort_data_table(Data_Table_object, "STATION")

                                                    If Data_Table_object.Rows.Count > 0 Then

                                                        ReDim Preserve Line_new(Index_l - 1)
                                                        Line_new(Index_l - 1) = New Line
                                                        Line_new(Index_l - 1).StartPoint = New Point3d(Data_Table_object.Rows(0).Item("X1"), Data_Table_object.Rows(0).Item("Y1"), 0)
                                                        Line_new(Index_l - 1).EndPoint = New Point3d(Data_Table_object.Rows(0).Item("X"), Data_Table_object.Rows(0).Item("Y"), 0)
                                                        Line_new(Index_l - 1).Layer = "NO PLOT"

                                                        Index_l = Index_l + 1


                                                        Data_Table_information.Rows.Add()


                                                        Data_Table_information.Rows(Index_row).Item("X") = Data_Table_object.Rows(0).Item("X")
                                                        Data_Table_information.Rows(Index_row).Item("Y") = Data_Table_object.Rows(0).Item("Y")
                                                        Data_Table_information.Rows(Index_row).Item("Z") = 0
                                                        Data_Table_information.Rows(Index_row).Item("STATION") = Data_Table_object.Rows(0).Item("STATION")
                                                        Data_Table_information.Rows(Index_row).Item("OFFSET") = Data_Table_object.Rows(0).Item("OFFSET")





                                                        Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT") = Data_Table_object.Rows(0).Item("LEFT_RIGHT")




                                                        Data_Table_information.Rows(Index_row).Item("LAYER_NAME") = Data_Table_object.Rows(0).Item("LAYER_NAME")
                                                        If IsDBNull(Data_Table_object.Rows(0).Item("LAYER_DESCRIPTION")) = False Then
                                                            Data_Table_information.Rows(Index_row).Item("LAYER_DESCRIPTION") = Data_Table_object.Rows(0).Item("LAYER_DESCRIPTION")
                                                        End If


                                                        Dim mlED As New MLeader

                                                        mlED = Creaza_Mleader_nou_fara_UCS_transform_CU_btrecord_AND_TRANS(BTrecord, Trans1, New Point3d(Data_Table_information.Rows(Index_row).Item("X"), Data_Table_information.Rows(Index_row).Item("Y"), 0),
                                                                                                               "East = " & Round(Data_Table_information.Rows(Index_row).Item("X"), 3) & vbCrLf &
                                                                                                               "North = " & Round(Data_Table_information.Rows(Index_row).Item("Y"), 3) & vbCrLf &
                                                                                                               "Station = " & Data_Table_information.Rows(Index_row).Item("STATION") & vbCrLf &
                                                                                                               "Offset = " & Round(Data_Table_information.Rows(Index_row).Item("OFFSET"), 3) & Data_Table_information.Rows(Index_row).Item("LEFT_RIGHT"), 1, 0.2, 0.5, 5, 5)



                                                        mlED.Layer = "NO PLOT"


                                                        Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Object_ID_OD, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                            If IsNothing(Records1) = False Then
                                                                If Records1.Count > 0 Then
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

                                                                            If Data_Table_information.Columns.Contains(Nume_field) = False Then
                                                                                Data_Table_information.Columns.Add(Nume_field, GetType(String))
                                                                            End If
                                                                            If Not Replace(Valoare_field, " ", "") = "" Then
                                                                                Data_Table_information.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                            End If
                                                                        Next
                                                                    Next
                                                                End If
                                                            End If
                                                        End Using



                                                        Index_row = Index_row + 1


                                                    End If
                                                End If


                                            End If

                                        End If
                                    End If


                                Next


                                If IsNothing(Line_new) = False Then
                                    If Line_new.Length > 0 Then
                                        For Jj = 0 To Line_new.Length - 1

                                            BTrecord.AppendEntity(Line_new(Jj))
                                            Trans1.AddNewlyCreatedDBObject(Line_new(Jj), True)
                                        Next
                                    End If
                                End If




                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                If Data_Table_information.Rows.Count > 0 Then
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                    For i = 0 To Data_Table_information.Columns.Count - 1
                        W1.Cells(1, i + 1).value2 = Data_Table_information.Columns(i).ColumnName
                    Next
                    Dim Rand_excel As Double = 2
                    For i = 0 To Data_Table_information.Rows.Count - 1
                        For j = 0 To Data_Table_information.Columns.Count - 1
                            If IsDBNull(Data_Table_information.Rows(i).Item(j)) = False Then
                                W1.Cells(Rand_excel, j + 1).value2 = Data_Table_information.Rows(i).Item(j)
                            End If
                        Next
                        Rand_excel = Rand_excel + 1
                    Next
                End If

                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Private Sub Button_point_to_Station_usa_Click(sender As Object, e As EventArgs) Handles Button_point_to_Station_usa.Click
        If Freeze_operations = False Then
            Try
                Try
                    If TextBox_east_intersection.Text = "" Then
                        MsgBox("Please specify the East COLUMN!")
                        Exit Sub
                    End If
                    If TextBox_north_intersection.Text = "" Then
                        MsgBox("Please specify the North COLUMN!")
                        Exit Sub
                    End If



                    If TextBox_station.Text = "" Then
                        MsgBox("Please specify the Station COLUMN!")
                        Exit Sub
                    End If

                    If TextBox_offset.Text = "" Then
                        MsgBox("Please specify the Offset COLUMN!")
                        Exit Sub
                    End If

                    Freeze_operations = True
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Using lock As DocumentLock = ThisDrawing.LockDocument
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                        Object_Prompt.SingleOnly = True
                        Rezultat1 = Editor1.GetSelection(Object_Prompt)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                        Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                        Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                        Dim Col_elev As String = TextBox_elevation_INTERSECTION.Text.ToUpper
                        Dim Col_Station As String = TextBox_station.Text.ToUpper
                        Dim start1 As Integer = CInt(TextBox_start.Text)
                        Dim end1 As Integer = CInt(TextBox_end.Text)
                        Dim Col_Point_number As String = TextBox_Point_name.Text.ToUpper
                        Dim Col_Descr As String = TextBox_description.Text.ToUpper
                        Dim Col_offset As String = TextBox_offset.Text.ToUpper
                        Dim Offset1 As Double = 200
                        If IsNumeric(TextBox_buffer_for_excelPT.Text) = True Then
                            Offset1 = CDbl(TextBox_buffer_for_excelPT.Text)

                        End If

                        If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            If IsNothing(Rezultat1) = False Then
                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                    Dim Data_table_3D_POLY As New System.Data.DataTable
                                    Data_table_3D_POLY.Columns.Add("X", GetType(Double))
                                    Data_table_3D_POLY.Columns.Add("Y", GetType(Double))
                                    Data_table_3D_POLY.Columns.Add("Z", GetType(Double))
                                    Dim indexdt As Double = 0
                                    Dim Poly3D As Polyline3d
                                    Dim Poly2D As Polyline

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(0)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                    If TypeOf Ent1 Is Polyline3d Or TypeOf Ent1 Is Polyline Then


                                        Poly3D = TryCast(Ent1, Polyline3d)
                                        Poly2D = TryCast(Ent1, Polyline)

                                        If IsNothing(Poly2D) = True Then
                                            Poly2D = New Polyline
                                            Dim Index2d As Double = 0
                                            For Each ObjId As ObjectId In Poly3D
                                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                                Index2d = Index2d + 1
                                            Next
                                            Poly2D.Elevation = 0
                                        End If

                                        Dim Layer1 As String
                                        Layer1 = "_00points"

                                        Creaza_layer(Layer1, 3, "created by point insertor", False)

                                        Dim Layer2 As String
                                        Layer2 = "NO PLOT"

                                        Creaza_layer(Layer1, 3, "NO PLOT", False)


                                        For i = start1 To end1

                                            Dim X As Double
                                            Dim Xstring As String = W1.Range(Col_east & i).Value

                                            Dim Y As Double
                                            Dim Ystring As String = W1.Range(Col_north & i).Value
                                            Dim Z As Double = 0
                                            Dim zstring As String = ""
                                            If Not TextBox_elevation_INTERSECTION.Text = "" Then
                                                zstring = W1.Range(Col_elev & i).Value
                                            End If

                                            If IsNumeric(Xstring) = True And IsNumeric(Ystring) = True Then
                                                X = CDbl(Xstring)
                                                Y = CDbl(Ystring)
                                                If IsNumeric(zstring) = True Then
                                                    Z = CDbl(zstring)
                                                End If





                                                Dim Text1 As New Autodesk.AutoCAD.DatabaseServices.DBText()
                                                Dim Text2 As New Autodesk.AutoCAD.DatabaseServices.DBText()
                                                Dim text3 As New Autodesk.AutoCAD.DatabaseServices.DBText()

                                                Dim PN As String = i
                                                If Not Col_Point_number = "" Then
                                                    PN = W1.Range(Col_Point_number & i).Value
                                                End If
                                                If Replace(PN, " ", "") = "" Then PN = i

                                                Text1.TextString = " " & PN
                                                Text1.Position = New Autodesk.AutoCAD.Geometry.Point3d(X, Y, Z)
                                                Text1.Height = 0.1
                                                Text1.Rotation = 7 * PI / 4
                                                Text1.Layer = Layer1
                                                BTrecord.AppendEntity(Text1)
                                                Trans1.AddNewlyCreatedDBObject(Text1, True)

                                                Dim Descriptie1 As String = "XXX"
                                                If Not Col_Descr = "" Then
                                                    Descriptie1 = W1.Range(Col_Descr & i).Value
                                                End If
                                                If Replace(PN, " ", "") = "" Then Descriptie1 = "XXX"


                                                Text2.TextString = Descriptie1
                                                Text2.Position = New Autodesk.AutoCAD.Geometry.Point3d(X, Y, Z)

                                                Text2.Height = 0.1
                                                Text2.Rotation = 0
                                                Text2.Layer = Layer1
                                                BTrecord.AppendEntity(Text2)
                                                Trans1.AddNewlyCreatedDBObject(Text2, True)




                                                text3.TextString = " " & Z




                                                text3.Position = New Autodesk.AutoCAD.Geometry.Point3d(X, Y, Z)

                                                text3.Height = 0.1
                                                text3.Rotation = PI / 4
                                                text3.Layer = Layer1
                                                BTrecord.AppendEntity(text3)
                                                Trans1.AddNewlyCreatedDBObject(text3, True)





                                                Dim Point_on_poly2d As New Point3d
                                                Dim Point_on_poly3d As New Point3d
                                                Point_on_poly2d = Poly2D.GetClosestPointTo(New Point3d(X, Y, 0), Vector3d.ZAxis, False)
                                                Dim Station1 As Double
                                                If IsNothing(Poly3D) = False Then
                                                    Point_on_poly3d = Poly3D.GetPointAtParameter(Poly2D.GetParameterAtPoint(Point_on_poly2d))
                                                    Station1 = Poly3D.GetDistAtPoint(Poly3D.GetClosestPointTo(Point_on_poly3d, Vector3d.ZAxis, False))
                                                ElseIf IsNothing(Poly2D) = False Then
                                                    Station1 = Poly2D.GetDistAtPoint(Poly2D.GetClosestPointTo(Point_on_poly2d, Vector3d.ZAxis, False))
                                                End If

                                                If IsNothing(Poly3D) = False Or IsNothing(Poly2D) = False Then

                                                    Dim Line1 As Line
                                                    Dim new_leader As New MLeader

                                                    If IsNothing(Poly3D) = False Then

                                                        Line1 = New Line(New Point3d(X, Y, Point_on_poly3d.Z), New Point3d(Point_on_poly3d.X, Point_on_poly3d.Y, Point_on_poly3d.Z))

                                                        If Line1.Length <= Offset1 Then

                                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly3d,
                                                                                                               "East = " & Round(Point_on_poly3d.X, 3) & vbCrLf &
                                                                                                               "North = " & Round(Point_on_poly3d.Y, 3) & vbCrLf &
                                                                                                               "Elev = " & Round(Point_on_poly3d.Z, 3) & vbCrLf &
                                                                                                               "Station = " & Get_chainage_feet_from_double(Station1, 0) _
                                                                                                               , 1, 0.2, 0.5, 5, 5)

                                                            W1.Range(Col_Station & i).Value = Round(Station1, 3)
                                                            W1.Range(Col_offset & i).Value = Round(Line1.Length, 3)
                                                        End If
                                                    ElseIf IsNothing(Poly2D) = False Then



                                                        Line1 = New Line(New Point3d(X, Y, 0), New Point3d(Point_on_poly2d.X, Point_on_poly2d.Y, 0))
                                                        If Line1.Length <= Offset1 Then
                                                            new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly2d,
                                                                                                               "East = " & Round(Point_on_poly2d.X, 3) & vbCrLf &
                                                                                                               "North = " & Round(Point_on_poly2d.Y, 3) & vbCrLf &
                                                                                                               "Station = " & Get_chainage_feet_from_double(Station1, 0) _
                                                                                                               , 1, 0.2, 0.5, 5, 5)
                                                            W1.Range(Col_Station & i).Value = Round(Station1, 3)
                                                            W1.Range(Col_offset & i).Value = Round(Line1.Length, 3)
                                                        End If



                                                    End If

                                                    new_leader.Layer = Layer2
                                                    Line1.Layer = Layer2
                                                    If Line1.Length <= Offset1 Then
                                                        BTrecord.AppendEntity(Line1)
                                                        Trans1.AddNewlyCreatedDBObject(Line1, True)
                                                    End If


                                                End If


                                            End If

                                        Next




                                    End If





                                    Trans1.Commit()
                                End Using
                                Editor1.Regen()
                            End If
                        End If
                    End Using

                    Freeze_operations = False
                    MsgBox("Done")
                    ThisDrawing.Editor.WriteMessage(vbLf & "Command:")


                Catch ex As Exception
                    Freeze_operations = False
                    MsgBox(ex.Message & vbCrLf & "Exception")
                End Try
                Freeze_operations = False


            Catch ex As System.SystemException
                Freeze_operations = False
                MsgBox(ex.Message & vbCrLf & "SystemException")
            End Try
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

    Private Sub Button_station_to_pointUSA_Click(sender As Object, e As EventArgs) Handles Button_station_to_pointUSA.Click

        If Freeze_operations = False Then
            Try
                Try
                    If TextBox_east_intersection.Text = "" Then
                        MsgBox("Please specify the East COLUMN!")
                        Exit Sub
                    End If
                    If TextBox_north_intersection.Text = "" Then
                        MsgBox("Please specify the North COLUMN!")
                        Exit Sub
                    End If

                    If TextBox_elevation_INTERSECTION.Text = "" Then
                        MsgBox("Please specify the ELEVATION COLUMN!")
                        Exit Sub
                    End If


                    If TextBox_station.Text = "" Then
                        MsgBox("Please specify the Station COLUMN!")
                        Exit Sub
                    End If


                    Freeze_operations = True
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Using lock As DocumentLock = ThisDrawing.LockDocument
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select 3D centerline:"
                        Object_Prompt.SingleOnly = True
                        Rezultat1 = Editor1.GetSelection(Object_Prompt)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                        Dim Col_east As String = TextBox_east_intersection.Text.ToUpper
                        Dim Col_north As String = TextBox_north_intersection.Text.ToUpper
                        Dim Col_elev As String = TextBox_elevation_INTERSECTION.Text.ToUpper
                        Dim Col_Station As String = TextBox_station.Text.ToUpper
                        Dim start1 As Integer = CInt(TextBox_start.Text)
                        Dim end1 As Integer = CInt(TextBox_end.Text)


                        If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            If IsNothing(Rezultat1) = False Then
                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                                    Dim Poly3D As Polyline3d
                                    Dim Poly2D As Polyline

                                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(0)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                    Dim Len1 As Double = -1
                                    If TypeOf Ent1 Is Polyline3d Or TypeOf Ent1 Is Polyline Then


                                        Poly3D = TryCast(Ent1, Polyline3d)
                                        Poly2D = TryCast(Ent1, Polyline)

                                        If IsNothing(Poly2D) = True Then
                                            Len1 = Poly3D.Length
                                            Poly2D = New Polyline
                                            Dim Index2d As Double = 0
                                            For Each ObjId As ObjectId In Poly3D
                                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, OpenMode.ForRead)
                                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                                Index2d = Index2d + 1
                                            Next
                                            Poly2D.Elevation = 0
                                        Else
                                            Len1 = Poly2D.Length
                                        End If


                                        Dim No_plot As String
                                        No_plot = "NO PLOT"

                                        Creaza_layer(No_plot, 40, No_plot, False)


                                        For i = start1 To end1

                                            Dim Station As Double = W1.Range(Col_Station & i).Value2

                                            If IsNumeric(Station) = True And Len1 >= Station Then

                                                Dim Point_on_poly As New Point3d

                                                Point_on_poly = Poly2D.GetPointAtDist(Station)

                                                If IsNothing(Poly3D) = False Then
                                                    Point_on_poly = Poly3D.GetPointAtDist(Station)
                                                End If

                                                If IsNothing(Poly3D) = False Or IsNothing(Poly2D) = False Then


                                                    Dim new_leader As New MLeader
                                                    new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly,
                                                                                                           "East = " & Round(Point_on_poly.X, 3) & vbCrLf &
                                                                                                           "North = " & Round(Point_on_poly.Y, 3) & vbCrLf &
                                                                                                           "Elev = " & Round(Point_on_poly.Z, 3) & vbCrLf &
                                                                                                           "Station = " & Get_chainage_feet_from_double(Station, 0) _
                                                                                                           , 1, 0.2, 0.5, 5, 5)
                                                    W1.Range(Col_east & i).Value = Round(Point_on_poly.X, 3)
                                                    W1.Range(Col_north & i).Value = Round(Point_on_poly.Y, 3)
                                                    W1.Range(Col_elev & i).Value = Round(Point_on_poly.Z, 3)

                                                    new_leader.Layer = No_plot






                                                End If


                                            End If

                                        Next




                                    End If





                                    Trans1.Commit()
                                End Using
                                Editor1.Regen()
                            End If
                        End If
                    End Using

                    Freeze_operations = False
                    MsgBox("Done")
                    ThisDrawing.Editor.WriteMessage(vbLf & "Command:")


                Catch ex As System.Exception
                    Freeze_operations = False
                    MsgBox(ex.Message & vbCrLf & "Exception")
                End Try

                Freeze_operations = False

            Catch ex As System.SystemException
                Freeze_operations = False
                MsgBox(ex.Message & vbCrLf & "SystemException")
            End Try
        End If

    End Sub

    Private Sub Button_calculate_2DCL_3Dxing_Click(sender As Object, e As EventArgs) Handles Button_cl_2d_crossing_3D.Click
        If Freeze_operations = False Then
            Try

                Dim Col_station As String = "STATION"

                Dim dt1 As New System.Data.DataTable
                dt1.Columns.Add(Col_station, GetType(Double))
                dt1.Columns.Add("DESCRIPTION", GetType(String))
                dt1.Columns.Add("X", GetType(Double))
                dt1.Columns.Add("Y", GetType(Double))
                dt1.Columns.Add("Z", GetType(Double))
                dt1.Columns.Add("LAYER_NAME", GetType(String))
                dt1.Columns.Add("LAYER_DESCRIPTION", GetType(String))


                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                Using lock As DocumentLock = ThisDrawing.LockDocument
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions

                    Object_Prompt.MessageForAdding = vbLf & "Select centerline:"


                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then
                            Freeze_operations = True
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Creaza_layer("NO PLOT", 40, "NO PLOT", False)
                                Dim Col_ObjId_CL As New ObjectIdCollection

                                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)



                                Dim Index_row As Double = 0

                                Dim ChainageCSF_colection As New DBObjectCollection
                                Dim CSF_colection As New DBObjectCollection

                                If TypeOf Ent1 Is Polyline Then
                                    Col_ObjId_CL.Add(Obj1.ObjectId)
                                End If

                                If Col_ObjId_CL.Count = 0 Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If



                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables

                                Dim Poly_2D_CL As Polyline = Trans1.GetObject(Col_ObjId_CL(0), OpenMode.ForRead)

                                For Each ObjID In BTrecord
                                    Dim Poly_3D_int As Polyline3d = TryCast(Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline3d)

                                    If Not Poly_3D_int = Nothing Then
                                        Dim Poly_2D_int As New Polyline
                                        Dim Index2d As Double = 0
                                        For Each ObjId1 As ObjectId In Poly_3D_int
                                            Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                            Poly_2D_int.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                            Index2d = Index2d + 1
                                        Next

                                        Poly_2D_int.Elevation = Poly_2D_CL.Elevation


                                        Dim Col_int As New Point3dCollection

                                        Col_int = Intersect_on_both_operands(Poly_2D_int, Poly_2D_CL)



                                        If Col_int.Count > 0 Then

                                            For index = 0 To Col_int.Count - 1
                                                Dim Point_on_poly_cl As New Point3d
                                                Dim Point_on_poly2d As New Point3d
                                                Dim Point_on_poly3d As New Point3d
                                                Dim sta1 As Double = 0
                                                Dim ChainageCSF0 As Double = 0

                                                Point_on_poly_cl = Poly_2D_CL.GetClosestPointTo(Col_int(index), Vector3d.ZAxis, False)
                                                Point_on_poly2d = Poly_2D_int.GetClosestPointTo(Col_int(index), Vector3d.ZAxis, False)

                                                Dim par_cl As Double = Poly_2D_CL.GetParameterAtPoint(Point_on_poly_cl)
                                                Dim par1 As Double = Poly_2D_int.GetParameterAtPoint(Point_on_poly2d)

                                                If par1 > Poly_3D_int.EndParam Then
                                                    par1 = Poly_3D_int.EndParam
                                                End If
                                                Point_on_poly3d = Poly_3D_int.GetPointAtParameter(par1)



                                                sta1 = Poly_2D_CL.GetDistanceAtParameter(par_cl)




                                                dt1.Rows.Add()
                                                dt1.Rows(Index_row).Item("X") = Round(Point_on_poly_cl.X, 3)
                                                dt1.Rows(Index_row).Item("Y") = Round(Point_on_poly_cl.Y, 3)
                                                dt1.Rows(Index_row).Item("Z") = Round(Point_on_poly3d.Z, 3)


                                                dt1.Rows(Index_row).Item(Col_station) = Round(sta1, 3)









                                                Dim Valoare As String = Poly_3D_int.Layer
                                                Dim Valoare1 As String = ""
                                                If IsNothing(Data_table_Layers) = False Then
                                                    If Data_table_Layers.Rows.Count > 0 Then
                                                        For J = 0 To Data_table_Layers.Rows.Count - 1
                                                            If IsDBNull(Data_table_Layers.Rows(J).Item("NAME")) = False And IsDBNull(Data_table_Layers.Rows(J).Item("DESCRIPTION")) = False Then
                                                                If Data_table_Layers.Rows(J).Item("NAME").ToString.ToUpper = Valoare.ToUpper Then
                                                                    Valoare1 = Data_table_Layers.Rows(J).Item("DESCRIPTION")
                                                                    Exit For
                                                                End If

                                                            End If
                                                        Next
                                                    End If
                                                End If
                                                dt1.Rows(Index_row).Item("LAYER_NAME") = Poly_3D_int.Layer
                                                If Not Valoare1 = "" Then
                                                    dt1.Rows(Index_row).Item("LAYER_DESCRIPTION") = Valoare1
                                                End If


                                                If CheckBox_Object_data.Checked = True Then
                                                    Dim Id1 As ObjectId = Poly_3D_int.ObjectId
                                                    Using Records1 As Autodesk.Gis.Map.ObjectData.Records = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                                        If IsNothing(Records1) = False Then
                                                            If Records1.Count > 0 Then
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

                                                                        If dt1.Columns.Contains(Nume_field) = False Then
                                                                            dt1.Columns.Add(Nume_field, GetType(String))
                                                                        End If
                                                                        If Not Replace(Valoare_field, " ", "") = "" Then
                                                                            dt1.Rows(Index_row).Item(Nume_field) = Valoare_field
                                                                        End If
                                                                    Next
                                                                Next
                                                            End If
                                                        End If
                                                    End Using

                                                End If

                                                Index_row = Index_row + 1

                                                Dim new_leader As New MLeader




                                                new_leader = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly3d,
                                                                                                           "East = " & Round(Point_on_poly3d.X, 3) & vbCrLf &
                                                                                                           "North = " & Round(Point_on_poly3d.Y, 3) & vbCrLf &
                                                                                                           "Elev = " & Round(Point_on_poly3d.Z, 3) & vbCrLf &
                                                                                                           "Station Grid = " & Get_chainage_from_double(sta1, 3), 1, 0.2, 0.5, 5, 5)
                                                new_leader.Layer = "NO PLOT"

                                            Next
                                        End If






                                    End If
                                Next






                                Editor1.Regen()
                                Trans1.Commit()
                            End Using
                        End If
                    End If
                End Using

                Transfer_datatable_to_new_excel_spreadsheet(dt1)


                MsgBox("Done")
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

                Freeze_operations = False
            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

End Class