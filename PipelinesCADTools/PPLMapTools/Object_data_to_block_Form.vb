Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Object_data_to_block_Form
    Dim Colectie1 As New Specialized.StringCollection

    Dim Index_combo As Integer = 1
    Dim Colectie_nume_OD As New Specialized.StringCollection

    Private Sub Object_data_to_block_Form_Click(sender As Object, e As EventArgs) Handles Me.Click
        Incarca_existing_layers_to_combobox(ComboBox_layers)
        If ComboBox_layers.Items.Count > 0 Then
            If ComboBox_layers.Items.Contains("TEXT") = True Then
                ComboBox_layers.SelectedIndex = ComboBox_layers.Items.IndexOf("TEXT")
            Else
                ComboBox_layers.SelectedIndex = 0
            End If
        End If
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        Button_load_attributes.Visible = False
        Panel_od.Controls.Clear()
        Colectie_nume_OD.Clear()
    End Sub

    Private Sub Object_data_to_block_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Incarca_existing_layers_to_combobox(ComboBox_layers)
        If ComboBox_layers.Items.Count > 0 Then
            If ComboBox_layers.Items.Contains("TEXT") = True Then
                ComboBox_layers.SelectedIndex = ComboBox_layers.Items.IndexOf("TEXT")
            Else
                ComboBox_layers.SelectedIndex = 0
            End If
        End If
        Incarca_existing_Blocks_with_attributes_to_combobox(ComboBox_blocks)
        Button_load_attributes.Visible = False
    End Sub

    Private Sub Button_read_OD_Click(sender As Object, e As EventArgs) Handles Button_read_OD.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select entity containing object data:")

            Rezultat1 = Editor1.GetEntity(Object_Prompt)

            If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                            Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables
                            Dim Id1 As ObjectId = Ent1.ObjectId
                            Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                            If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)
                                Dim Record1 As Autodesk.Gis.Map.ObjectData.Record
                                Panel_od.Controls.Clear()
                                Colectie_nume_OD = New Specialized.StringCollection



                                For Each Record1 In Records1
                                    Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                    Tabla1 = Tables1(Record1.TableName)

                                    Dim Combo1 As New Windows.Forms.ComboBox
                                    Combo1.Location = New Drawing.Point(108, 3)
                                    Combo1.Size = New Drawing.Size(150, 23)
                                    Combo1.Name = "COMBOBOX_" & Index_combo
                                    Panel_od.Controls.Add(Combo1)

                                    Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                    Field_defs1 = Tabla1.FieldDefinitions
                                    For i = 0 To Record1.Count - 1
                                        Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                        Field_def1 = Field_defs1(i)

                                        Combo1.Items.Add(Field_def1.Name)
                                        Colectie_nume_OD.Add(Field_def1.Name)
                                    Next
                                    If Combo1.Items.Contains("LINELIST") = True Then
                                        Combo1.SelectedIndex = Combo1.Items.IndexOf("LINELIST")
                                    End If
                                    Button_load_attributes.Visible = True

                                Next


                            End If



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
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub


    Private Sub Button_load_attributes_Click(sender As Object, e As EventArgs) Handles Button_load_attributes.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        Try

            Colectie1 = New Specialized.StringCollection
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecordBlock As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    If Not Replace(ComboBox_blocks.Text, " ", "") = "" Then

                        BTrecordBlock = Trans1.GetObject(BlockTable_data1(ComboBox_blocks.Text), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If IsNothing(BTrecordBlock) = False Then
                            If BTrecordBlock.HasAttributeDefinitions = True Then
                                Dim i As Integer = 0
                                Index_combo = 1
                                Panel_od.Controls.Clear()
                                If Colectie_nume_OD.Count > 0 Then
                                    For Each Id1 As ObjectId In BTrecordBlock
                                        Dim ent As Entity = TryCast(Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                                        If ent IsNot Nothing Then
                                            Dim attDefinition1 As AttributeDefinition = TryCast(ent, AttributeDefinition)
                                            If attDefinition1 IsNot Nothing Then


                                                Dim TextBox2 As New Windows.Forms.TextBox
                                                TextBox2.Location = New Drawing.Point(3, 5 + 26 * i)
                                                TextBox2.Size = New Drawing.Size(100, 23)
                                                TextBox2.Text = attDefinition1.Tag
                                                Panel_od.Controls.Add(TextBox2)

                                                Dim Combo2 As New Windows.Forms.ComboBox
                                                Combo2.Location = New Drawing.Point(108, 5 + 26 * i)
                                                Combo2.Size = New Drawing.Size(150, 23)
                                                For j = 0 To Colectie_nume_OD.Count - 1
                                                    Combo2.Items.Add(Colectie_nume_OD(j))
                                                Next

                                                Panel_od.Controls.Add(Combo2)

                                                If i < 7 Then
                                                    If i = 0 Then
                                                        Panel_od.Height = 112
                                                    Else
                                                        Panel_od.Height = Combo2.Top + Combo2.Height + 8
                                                    End If

                                                    If Panel_od.Top + Panel_od.Height + 8 > 139 Then
                                                        Button_read_OD.Top = Panel_od.Top + Panel_od.Height + 8
                                                    End If

                                                    Button_load_attributes.Top = Button_read_OD.Top
                                                    Button_insert_block.Top = Button_load_attributes.Top


                                                    Me.Height = Button_insert_block.Top + Button_insert_block.Height + 40
                                                End If

                                                i = i + 1

                                            End If
                                        End If
                                    Next
                                End If

                            End If
                        End If
                    End If

                End Using
            End Using


            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")

        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_insert_block_Click(sender As Object, e As EventArgs) Handles Button_insert_block.Click
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                ' Dim k As Double = 1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_poly_with_od As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Prompt_poly_OD As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Prompt_poly_OD.MessageForAdding = vbLf & "Select the polylines/points for which you want to insert the blocks:"

                    Prompt_poly_OD.SingleOnly = False
                    Rezultat_poly_with_od = Editor1.GetSelection(Prompt_poly_OD)


                    Dim Rezultat_poly_rotatie As Autodesk.AutoCAD.EditorInput.PromptEntityResult
                    Dim Prompt_poly_rotatie As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the polyline you want to align the blocks:")
                    Rezultat_poly_rotatie = Editor1.GetEntity(Prompt_poly_rotatie)

                    Dim Colectie_obiecte_de_sters As New DBObjectCollection

                    If Rezultat_poly_with_od.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        For i = 0 To Rezultat_poly_with_od.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject = Rezultat_poly_with_od.Value(i)
                            Dim Ent1 As Entity = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline Then
                                Dim Poly1 As Polyline = Ent1

                                Dim Colectie_atribut_block_names As New Specialized.StringCollection
                                Dim Colectie_atribut_poly_names As New Specialized.StringCollection
                                Dim Colectie_values_poly As New Specialized.StringCollection
                                Dim Colectie_values_block As New Specialized.StringCollection


                                Dim Id1 As ObjectId = Poly1.ObjectId
                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables


                                Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                                If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                    Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)

                                    For Each Record1 As Autodesk.Gis.Map.ObjectData.Record In Records1
                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                        Tabla1 = Tables1(Record1.TableName)

                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For j = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(j)
                                            If Not Replace(Record1(j).StrValue, " ", "") = "" Then
                                                Colectie_atribut_poly_names.Add(Field_def1.Name)
                                                Colectie_values_poly.Add(Record1(j).StrValue)
                                            End If
                                        Next
                                    Next
                                End If

                                For Each Control1 As Windows.Forms.Control In Panel_od.Controls
                                    If TypeOf Control1 Is Windows.Forms.TextBox Then
                                        Dim TextBox3 As Windows.Forms.TextBox = Control1
                                        If Not Replace(TextBox3.Text, " ", "") = "" Then
                                            For Each Control2 As Windows.Forms.Control In Panel_od.Controls
                                                If TypeOf Control2 Is Windows.Forms.ComboBox Then
                                                    Dim Combo3 As Windows.Forms.ComboBox = Control2
                                                    If Combo3.Top = TextBox3.Top And Not Replace(Combo3.Text, " ", "") = "" Then
                                                        If Colectie_atribut_poly_names.Contains(Combo3.Text) = True Then
                                                            Colectie_atribut_block_names.Add(TextBox3.Text)
                                                            If CheckBox_upper_case.Checked = True Then
                                                                Colectie_values_block.Add(Colectie_values_poly(Colectie_atribut_poly_names.IndexOf(Combo3.Text)).ToUpper)
                                                            Else
                                                                Colectie_values_block.Add(Colectie_values_poly(Colectie_atribut_poly_names.IndexOf(Combo3.Text)))
                                                            End If

                                                        End If
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next

                                Dim Rotatie_block As Double = 0

                                If Rezultat_poly_rotatie.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                                    Dim Ent2 As Entity = Rezultat_poly_rotatie.ObjectId.GetObject(OpenMode.ForRead)

                                    If TypeOf Ent2 Is Polyline Then
                                        Dim Poly2 As Polyline = Ent2
                                        Dim Point_on_poly As New Point3d
                                        Point_on_poly = Poly2.GetClosestPointTo(Rezultat_poly_rotatie.PickedPoint, Vector3d.ZAxis, True)

                                        Dim Param_m As Double = Poly2.GetParameterAtPoint(Point_on_poly)

                                        Dim Param_0 As Double
                                        Dim Param_1 As Double
                                        If Round(Param_m, 2) = Round(Param_m, 0) Then
                                            If Round(Param_m, 0) = 0 Then
                                                Param_0 = 0
                                                Param_1 = 1

                                            ElseIf Round(Param_m, 0) = Poly2.NumberOfVertices - 1 Then
                                                Param_0 = Poly2.NumberOfVertices - 2
                                                Param_1 = Poly2.NumberOfVertices - 1

                                            Else
                                                If Poly2.NumberOfVertices - 1 >= Round(Param_m, 0) + 1 Then
                                                    Param_0 = Round(Param_m, 0)
                                                    Param_1 = Round(Param_m, 0) + 1
                                                Else
                                                    If Round(Param_m, 0) - 1 >= 0 Then
                                                        Param_0 = Round(Param_m, 0) - 1
                                                        Param_1 = Round(Param_m, 0)
                                                    End If

                                                End If
                                            End If
                                        Else
                                            Param_0 = Floor(Param_m)
                                            Param_1 = Ceiling(Param_m)
                                        End If

                                        If Poly2.NumberOfVertices > 1 Then
                                            Rotatie_block = GET_Bearing_rad(Poly2.GetPoint3dAt(Param_0).X, Poly2.GetPoint3dAt(Param_0).Y, Poly2.GetPoint3dAt(Param_1).X, Poly2.GetPoint3dAt(Param_1).Y)
                                        End If

                                    End If


                                End If

                                If CheckBox_rotate_180.Checked = True Then Rotatie_block = Rotatie_block + PI

                                Dim BlockScale As Double = 1
                                If IsNumeric(TextBox_block_scale.Text) = True Then
                                    BlockScale = CDbl(TextBox_block_scale.Text)
                                End If

                                Dim X As Double = 0
                                Dim Y As Double = 0
                                For k = 0 To Poly1.NumberOfVertices - 1
                                    X = X + Poly1.GetPoint2dAt(k).X
                                    Y = Y + Poly1.GetPoint2dAt(k).Y
                                Next
                                X = X / Poly1.NumberOfVertices - 1
                                Y = Y / Poly1.NumberOfVertices - 1

                                If CheckBox_specify_each_point.Checked = True Then

                                    Dim view1 As ViewTableRecord
                                    Dim View_Table As ViewTable = Trans1.GetObject(ThisDrawing.Database.ViewTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                                    Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")
                                    If Tilemode1 = 0 Then
                                        Application.SetSystemVariable("TILEMODE", 1)
                                    End If

                                    If View_Table.Has("DAN") = False Then
                                        View_Table.UpgradeOpen()
                                        view1 = New ViewTableRecord
                                        view1.CenterPoint = New Point2d(0, 0)
                                        view1.Target = New Point3d(X, Y, 0)
                                        view1.ViewTwist = 0
                                        view1.Width = 0.5 * Poly1.Length
                                        view1.Height = 0.5 * Poly1.Length
                                        view1.Name = "DAN"
                                        View_Table.Add(view1)
                                        Trans1.AddNewlyCreatedDBObject(view1, True)
                                    Else
                                        view1 = View_Table("DAN").GetObject(OpenMode.ForWrite)
                                        view1.CenterPoint = New Point2d(0, 0)
                                        view1.Target = New Point3d(X, Y, 0)
                                        view1.ViewTwist = 0
                                        view1.Width = 0.5 * Poly1.Length
                                        view1.Height = 0.5 * Poly1.Length
                                    End If
                                    ThisDrawing.Editor.SetCurrentView(view1)


                                    Dim Pt_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult



                                    Dim Jig1 As New Jig_highlight_poly_Class

                                    Pt_rezult = Jig1.StartJig(Poly1)


                                    If Pt_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        X = Pt_rezult.Value.X
                                        Y = Pt_rezult.Value.Y
                                    End If



                                End If


                                Dim Block1 As BlockReference = InsertBlock_with_multiple_atributes("", ComboBox_blocks.Text, New Point3d(X, Y, 0), BlockScale, BTrecord, ComboBox_layers.Text, Colectie_atribut_block_names, Colectie_values_block)
                                Block1.Rotation = Rotatie_block




                            End If

                            If TypeOf Ent1 Is DBPoint Then
                                Dim Point1 As DBPoint = Ent1

                                Dim Colectie_atribut_block_names As New Specialized.StringCollection
                                Dim Colectie_atribut_poly_names As New Specialized.StringCollection
                                Dim Colectie_values_poly As New Specialized.StringCollection
                                Dim Colectie_values_block As New Specialized.StringCollection


                                Dim Id1 As ObjectId = Point1.ObjectId
                                Dim Tables1 As Autodesk.Gis.Map.ObjectData.Tables
                                Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables


                                Dim Records1 As Autodesk.Gis.Map.ObjectData.Records
                                If Not Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False).Count = 0 Then
                                    Records1 = Tables1.GetObjectRecords(Convert.ToInt32(0), Id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, False)

                                    For Each Record1 As Autodesk.Gis.Map.ObjectData.Record In Records1
                                        Dim Tabla1 As Autodesk.Gis.Map.ObjectData.Table
                                        Tabla1 = Tables1(Record1.TableName)

                                        Dim Field_defs1 As Autodesk.Gis.Map.ObjectData.FieldDefinitions
                                        Field_defs1 = Tabla1.FieldDefinitions
                                        For j = 0 To Record1.Count - 1
                                            Dim Field_def1 As Autodesk.Gis.Map.ObjectData.FieldDefinition
                                            Field_def1 = Field_defs1(j)
                                            If Not Replace(Record1(j).StrValue, " ", "") = "" Then
                                                Colectie_atribut_poly_names.Add(Field_def1.Name)
                                                Colectie_values_poly.Add(Record1(j).StrValue)
                                            End If
                                        Next
                                    Next
                                End If

                                For Each Control1 As Windows.Forms.Control In Panel_od.Controls
                                    If TypeOf Control1 Is Windows.Forms.TextBox Then
                                        Dim TextBox3 As Windows.Forms.TextBox = Control1
                                        If Not Replace(TextBox3.Text, " ", "") = "" Then
                                            For Each Control2 As Windows.Forms.Control In Panel_od.Controls
                                                If TypeOf Control2 Is Windows.Forms.ComboBox Then
                                                    Dim Combo3 As Windows.Forms.ComboBox = Control2
                                                    If Combo3.Top = TextBox3.Top And Not Replace(Combo3.Text, " ", "") = "" Then
                                                        If Colectie_atribut_poly_names.Contains(Combo3.Text) = True Then
                                                            Colectie_atribut_block_names.Add(TextBox3.Text)
                                                            If CheckBox_upper_case.Checked = True Then
                                                                Colectie_values_block.Add(Colectie_values_poly(Colectie_atribut_poly_names.IndexOf(Combo3.Text)).ToUpper)
                                                            Else
                                                                Colectie_values_block.Add(Colectie_values_poly(Colectie_atribut_poly_names.IndexOf(Combo3.Text)))
                                                            End If

                                                        End If
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next

                                Dim Rotatie_block As Double = 0

                                If Rezultat_poly_rotatie.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                                    Dim Ent2 As Entity = Rezultat_poly_rotatie.ObjectId.GetObject(OpenMode.ForRead)

                                    If TypeOf Ent2 Is Polyline Then
                                        Dim Poly2 As Polyline = Ent2
                                        Dim Point_on_poly As New Point3d
                                        Point_on_poly = Poly2.GetClosestPointTo(Rezultat_poly_rotatie.PickedPoint, Vector3d.ZAxis, True)

                                        Dim Param_m As Double = Poly2.GetParameterAtPoint(Point_on_poly)

                                        Dim Param_0 As Double
                                        Dim Param_1 As Double
                                        If Round(Param_m, 2) = Round(Param_m, 0) Then
                                            If Round(Param_m, 0) = 0 Then
                                                Param_0 = 0
                                                Param_1 = 1

                                            ElseIf Round(Param_m, 0) = Poly2.NumberOfVertices - 1 Then
                                                Param_0 = Poly2.NumberOfVertices - 2
                                                Param_1 = Poly2.NumberOfVertices - 1

                                            Else
                                                If Poly2.NumberOfVertices - 1 >= Round(Param_m, 0) + 1 Then
                                                    Param_0 = Round(Param_m, 0)
                                                    Param_1 = Round(Param_m, 0) + 1
                                                Else
                                                    If Round(Param_m, 0) - 1 >= 0 Then
                                                        Param_0 = Round(Param_m, 0) - 1
                                                        Param_1 = Round(Param_m, 0)
                                                    End If

                                                End If
                                            End If
                                        Else
                                            Param_0 = Floor(Param_m)
                                            Param_1 = Ceiling(Param_m)
                                        End If

                                        If Poly2.NumberOfVertices > 1 Then
                                            Rotatie_block = GET_Bearing_rad(Poly2.GetPoint3dAt(Param_0).X, Poly2.GetPoint3dAt(Param_0).Y, Poly2.GetPoint3dAt(Param_1).X, Poly2.GetPoint3dAt(Param_1).Y)
                                        End If

                                    End If


                                End If

                                If CheckBox_rotate_180.Checked = True Then Rotatie_block = Rotatie_block + PI

                                Dim BlockScale As Double = 1
                                If IsNumeric(TextBox_block_scale.Text) = True Then
                                    BlockScale = CDbl(TextBox_block_scale.Text)
                                End If

                                Dim X As Double = Point1.Position.X
                                Dim Y As Double = Point1.Position.Y



                                


                                Dim Block1 As BlockReference = InsertBlock_with_multiple_atributes("", ComboBox_blocks.Text, New Point3d(X, Y, 0), BlockScale, BTrecord, ComboBox_layers.Text, Colectie_atribut_block_names, Colectie_values_block)
                                Block1.Rotation = Rotatie_block




                            End If
                        Next
                    End If








                    Trans1.Commit()
                    ' asta e de la tranzactie

                End Using


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                    Dim BlocktableRec1 As BlockTableRecord = BlockTable_data1.Item(ComboBox_blocks.Text).GetObject(OpenMode.ForWrite)
                    SynchronizeAttributes_db_diferit(BlocktableRec1, Trans1)
                    Trans1.Commit()
                End Using


                ' Dim k As Double = 1







                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox("Done")


                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                ' asta e de la lock
            End Using


            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub
End Class