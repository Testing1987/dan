Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Windows.Forms

Public Class Overlap_form
    Dim Freeze_operations As Boolean = False
    Dim Empty_array() As ObjectId
    Dim Ultimul_top As Double
    Dim Spacing As Double
    Private Sub Overlap_form_Load(sender As Object, e As EventArgs) Handles Me.Load, Button_refresh_layers.Click
        Incarca_existing_layers_to_combobox(ComboBox_layer1)
        Incarca_existing_layers_to_combobox(ComboBox_layer3)
        Incarca_existing_layers_to_combobox(ComboBox_layer2)
        Ultimul_top = ComboBox_layer3.Top
        Spacing = ComboBox_layer3.Top - ComboBox_layer2.Top

        Dim Del1() As Windows.Forms.Control
        Dim Index1 As Integer = 0
        For Each control1 As Windows.Forms.Control In Panel_layers.Controls
            If control1.Top > ComboBox_layer3.Top Then
                ReDim Preserve Del1(Index1)
                Del1(Index1) = control1
                Index1 = Index1 + 1
            End If
        Next

        If IsNothing(Del1) = False Then
            For i = 0 To Del1.Length - 1
                Panel_layers.Controls.Remove(Del1(i))
            Next
        End If


        Me.Refresh()

    End Sub
    Private Sub Button_add_new_combobox_Click(sender As Object, e As EventArgs) Handles Button_add_new_combobox.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try


                Dim Combo_1 As New Windows.Forms.ComboBox
                Combo_1.Size = ComboBox_layer3.Size
                Combo_1.Left = ComboBox_layer3.Left
                Combo_1.Top = Ultimul_top + Spacing
                Combo_1.DropDownStyle = ComboBox_layer3.DropDownStyle
                Combo_1.Font = ComboBox_layer3.Font
                Combo_1.ForeColor = ComboBox_layer3.ForeColor
                Combo_1.BackColor = ComboBox_layer3.BackColor

                Panel_layers.Controls.Add(Combo_1)

                Incarca_existing_layers_to_combobox(Combo_1)
                Ultimul_top = Combo_1.Top

            Catch ex As System.SystemException

                MsgBox(ex.Message)

            End Try
            Freeze_operations = False


        End If
    End Sub


    Private Sub Button_analise_gaps_Click(sender As Object, e As EventArgs) Handles Button_analise_gaps.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim No_plot As String = "NO PLOT"
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Colectie_poly1 As New DBObjectCollection
                    Dim Colectie_poly2 As New DBObjectCollection
                    Creaza_layer(No_plot, 40, "no plot", False)

                    Dim Lista1 As New Specialized.StringCollection
                    For Each Combo1 As Windows.Forms.Control In Panel_layers.Controls
                        If TypeOf Combo1 Is Windows.Forms.ComboBox Then
                            Dim Text1 As String = Combo1.Text
                            If Not Text1 = "" Then
                                If Lista1.Contains(Text1) = False Then
                                    Lista1.Add(Text1)
                                End If
                            End If
                        End If
                    Next
                    If Lista1.Count > 0 Then
                        For Each O_id In BTrecord
                            Dim Ent1 As Entity = Trans1.GetObject(O_id, OpenMode.ForRead)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim Poly1 As Polyline = Ent1
                                If Lista1.Contains(Poly1.Layer) = True Then
                                    Poly1.UpgradeOpen()
                                    Poly1.ColorIndex = 7
                                    If Poly1.Closed = True Then
                                        Colectie_poly1.Add(Poly1)
                                    Else
                                        Colectie_poly2.Add(Poly1)
                                    End If
                                End If


                            End If
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Region And Ent1.Layer = No_plot Then
                                Dim Region1 As Region = Ent1
                                Region1.UpgradeOpen()
                                Region1.Erase()
                            End If
                        Next



                        Dim RegionCol_union As DBObjectCollection = New DBObjectCollection()


                        If Colectie_poly1.Count > 0 Then
                            RegionCol_union = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Colectie_poly1)
                        End If

                        If RegionCol_union.Count > 1 Then
                            Dim Region1 As Region = RegionCol_union(0)
                            Region1.ColorIndex = 5
                            Region1.Layer = No_plot
                            BTrecord.AppendEntity(Region1)
                            Trans1.AddNewlyCreatedDBObject(Region1, True)
                            Try
                                For i = 1 To RegionCol_union.Count - 1
                                    Dim Region2 As Region = RegionCol_union(i)
                                    Region2.ColorIndex = 5
                                    Region1.BooleanOperation(BooleanOperationType.BoolUnite, Region2)
                                Next
                            Catch ex As Exception

                            End Try

                        End If


                        If Colectie_poly2.Count > 0 Then

                            For Each poly2 As Polyline In Colectie_poly2
                                Dim Center2 As New Point3d
                                Center2 = poly2.GetPointAtParameter(0)
                                Dim Cerc2 As New Circle(Center2, Vector3d.ZAxis, 10)
                                Cerc2.Layer = No_plot
                                BTrecord.AppendEntity(Cerc2)
                                Trans1.AddNewlyCreatedDBObject(Cerc2, True)


                            Next
                            If Colectie_poly2.Count = 1 Then
                                MsgBox("there is (1) not-closed polyline")
                            Else
                                MsgBox("there are (" & Colectie_poly2.Count & ") not-closed polylines")
                            End If

                            Editor1.WriteMessage(vbLf & Colectie_poly2.Count.ToString & " not-closed")
                        End If


                        Trans1.Commit()
                    End If
                End Using
            End Using




        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_analise_overlapp_Click(sender As Object, e As EventArgs) Handles Button_analise_overlapp.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim No_plot As String = "NO PLOT"
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Colectie_poly1 As New DBObjectCollection
                    Dim Colectie_poly2 As New DBObjectCollection
                    Creaza_layer(No_plot, 40, "no plot", False)
                    Dim Lista1 As New Specialized.StringCollection
                    For Each Combo1 As Windows.Forms.Control In Panel_layers.Controls
                        If TypeOf Combo1 Is Windows.Forms.ComboBox Then
                            Dim Text1 As String = Combo1.Text
                            If Not Text1 = "" Then
                                If Lista1.Contains(Text1) = False Then
                                    Lista1.Add(Text1)
                                End If
                            End If
                        End If
                    Next
                    If Lista1.Count > 0 Then
                        For Each O_id In BTrecord
                            Dim Ent1 As Entity = Trans1.GetObject(O_id, OpenMode.ForRead)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim Poly1 As Polyline = Ent1
                                If Lista1.Contains(Poly1.Layer) = True Then
                                    Poly1.UpgradeOpen()
                                    Poly1.ColorIndex = 7
                                    If Poly1.Closed = True Then
                                        Colectie_poly1.Add(Poly1)
                                    Else
                                        Colectie_poly2.Add(Poly1)
                                    End If
                                End If
                            End If
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Region And Ent1.Layer = No_plot Then
                                Dim Region1 As Region = Ent1
                                Region1.UpgradeOpen()
                                Region1.Erase()
                            End If
                        Next


                        Dim RegionCol_int As DBObjectCollection = New DBObjectCollection()

                        If Colectie_poly1.Count > 0 Then
                            RegionCol_int = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Colectie_poly1)
                        End If

                        If RegionCol_int.Count > 1 Then

                            For i = 0 To RegionCol_int.Count - 2
                                Dim Region1 As Region = RegionCol_int(i).Clone

                                For j = i + 1 To RegionCol_int.Count - 1
                                    Dim Region2 As Region = RegionCol_int(j).Clone
                                    Try
                                        Region1.BooleanOperation(BooleanOperationType.BoolIntersect, Region2)
                                        Region1.ColorIndex = 1
                                        Region1.Layer = No_plot
                                        BTrecord.AppendEntity(Region1)
                                        Trans1.AddNewlyCreatedDBObject(Region1, True)
                                        If Region1.Area < 0.001 Then
                                            Region1.Erase()
                                        End If

                                        Region1 = New Region
                                        Region1 = RegionCol_int(i).Clone

                                        Region2.Dispose()

                                    Catch ex As Exception

                                    End Try

                                Next
                            Next




                        End If


                        If Colectie_poly2.Count > 0 Then

                            For Each poly2 As Polyline In Colectie_poly2
                                Dim Center2 As New Point3d
                                Center2 = poly2.GetPointAtParameter(0)
                                Dim Cerc2 As New Circle(Center2, Vector3d.ZAxis, 10)
                                Cerc2.Layer = No_plot
                                BTrecord.AppendEntity(Cerc2)
                                Trans1.AddNewlyCreatedDBObject(Cerc2, True)


                            Next
                            If Colectie_poly2.Count = 1 Then
                                MsgBox("there is (1) not-closed polyline")
                            Else
                                MsgBox("there are (" & Colectie_poly2.Count & ") not-closed polylines")
                            End If

                            Editor1.WriteMessage(vbLf & Colectie_poly2.Count.ToString & " not-closed")
                        End If


                        Trans1.Commit()
                    End If
                End Using
            End Using




        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_dimension_To_CL_Click(sender As Object, e As EventArgs) Handles Button_dimension_To_CL.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim No_plot As String = "NO PLOT"
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")

                Object_Prompt.SetRejectMessage(vbLf & "Please select a lightweight polyline")
                Object_Prompt.AddAllowedClass(GetType(Polyline), True)


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


                                Dim PolyCL As Polyline = TryCast(Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Polyline)



                                If Not PolyCL.Elevation = 0 Then
                                    Freeze_operations = False
                                    MsgBox("CL Polyline is not at elevation 0")
                                    Exit Sub

                                End If


                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                Creaza_layer(No_plot, 40, "no plot", False)

                                Dim Lista1 As New Specialized.StringCollection
                                For Each Combo1 As Windows.Forms.Control In Panel_layers.Controls
                                    If TypeOf Combo1 Is Windows.Forms.ComboBox Then
                                        Dim Text1 As String = Combo1.Text
                                        If Not Text1 = "" Then
                                            If Lista1.Contains(Text1) = False Then
                                                Lista1.Add(Text1)
                                            End If
                                        End If
                                    End If
                                Next
                                If Lista1.Count > 0 Then
                                    For Each objID As ObjectId In BTrecord
                                        Dim Ent1 As Entity = Trans1.GetObject(objID, OpenMode.ForRead)
                                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.RotatedDimension And Ent1.Layer = No_plot Then
                                            Ent1.UpgradeOpen()
                                            Ent1.Erase()
                                        End If
                                    Next


                                    For Each objID As ObjectId In BTrecord
                                        Dim Ent1 As Entity = Trans1.GetObject(objID, OpenMode.ForRead)

                                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline And Not objID = PolyCL.ObjectId Then
                                            Dim Poly1 As Polyline = Ent1
                                            If Lista1.Contains(Poly1.Layer) = True Then
                                                Poly1.UpgradeOpen()
                                                Poly1.ColorIndex = 7
                                                For i = 0 To Poly1.NumberOfVertices - 2
                                                    Dim Mid_point_poly1 As New Point3d
                                                    Mid_point_poly1 = Poly1.GetPointAtParameter(i + 0.5)
                                                    Dim Bearing1 As Double = Round(GET_Bearing_rad(Poly1.GetPointAtParameter(i).X, Poly1.GetPointAtParameter(i).Y, Poly1.GetPointAtParameter(i + 1).X, Poly1.GetPointAtParameter(i + 1).Y), 4)
                                                    Dim Pt_on_poly As New Point3d
                                                    Pt_on_poly = PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, True)
                                                    Dim Paramcl As Double = PolyCL.GetParameterAtPoint(Pt_on_poly)

                                                    Dim PT_i As New Point3d
                                                    PT_i = PolyCL.GetPointAtParameter(Floor(Paramcl))

                                                    Dim PT_i_1 As New Point3d
                                                    PT_i_1 = PolyCL.GetPointAtParameter(Ceiling(Paramcl))

                                                    If Floor(Paramcl) = Ceiling(Paramcl) Then
                                                        If Floor(Paramcl) >= 1 And Floor(Paramcl) < PolyCL.NumberOfVertices - 1 Then
                                                            PT_i_1 = PolyCL.GetPointAtParameter(Floor(Paramcl) + 1)
                                                        Else
                                                            PT_i = PolyCL.GetPointAtParameter(0)
                                                            PT_i_1 = PolyCL.GetPointAtParameter(1)
                                                        End If
                                                    End If




                                                    Dim Bearing_cl As Double = Round(GET_Bearing_rad(PT_i.X, PT_i.Y, PT_i_1.X, PT_i_1.Y), 4)

                                                    Do Until Bearing_cl < Round(PI, 4)
                                                        If Bearing_cl >= Round(PI, 4) Then
                                                            Bearing_cl = Round(Bearing_cl - Round(PI, 4), 4)
                                                        End If
                                                    Loop

                                                    Do Until Bearing1 < Round(PI, 4)
                                                        If Bearing1 >= Round(PI, 4) Then
                                                            Bearing1 = Round(Bearing1 - Round(PI, 4), 4)
                                                        End If
                                                    Loop

                                                    If Bearing_cl = Bearing1 Or Bearing_cl + Round(PI, 4) = Bearing1 Or Bearing_cl - Round(PI, 4) = Bearing1 Then
                                                        Dim Line1 As New Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, True))
                                                        Line1.Layer = No_plot
                                                        Dim Offset As Double = Line1.Length
                                                        Dim a As Double = Round(Offset, 3)
                                                        Dim b As Double = Round(Offset, 0)
                                                        If Not a = b Then
                                                            Dim Dimension1 As New RotatedDimension
                                                            Dimension1.Layer = No_plot
                                                            Dimension1.XLine1Point = Line1.StartPoint
                                                            Dimension1.XLine2Point = Line1.EndPoint
                                                            Dimension1.Rotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                            Dimension1.DimLinePoint = Line1.StartPoint
                                                            Dimension1.UsingDefaultTextPosition = True
                                                            Dimension1.TextAttachment = AttachmentPoint.MiddleCenter
                                                            Dimension1.TextRotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                            Dimension1.Dimasz = 2 'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                            Dimension1.Dimdec = 4 'Sets the number of decimal places displayed for the primary units of a dimension
                                                            Dimension1.Dimtxt = 4 'Specifies the height of dimension text, unless the current text style has a fixed height
                                                            add_extra_param_to_dim(Dimension1, ThisDrawing)
                                                            BTrecord.AppendEntity(Dimension1)
                                                            Trans1.AddNewlyCreatedDBObject(Dimension1, True)
                                                        End If
                                                    Else


                                                        Dim Line1 As New Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, True))
                                                        Line1.Layer = No_plot

                                                        Dim L1 As Double = Line1.Length

                                                        Dim a As Double = Round((Bearing1 + Round(PI, 4) / 2), 4)
                                                        Dim b As Double = Round(Bearing_cl, 4)
                                                        Dim c As Double = Round((Bearing1 - Round(PI, 4) / 2), 4)
                                                        Dim d As Double = Round(Bearing1, 4)
                                                        Dim Add_L As Boolean = False

                                                        If Not a = b And Not c = b And Not d = b Then
                                                            Add_L = True
                                                        End If

                                                        If Round(L1, 0) = Round(L1, 3) Then
                                                            Add_L = False
                                                        End If


                                                        If Add_L = True Then
                                                            Dim Dimension1 As New RotatedDimension
                                                            Dimension1.Layer = No_plot
                                                            Dimension1.XLine1Point = Line1.StartPoint
                                                            Dimension1.XLine2Point = Line1.EndPoint
                                                            Dimension1.Rotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                            Dimension1.DimLinePoint = Line1.StartPoint
                                                            Dimension1.UsingDefaultTextPosition = True
                                                            Dimension1.TextAttachment = AttachmentPoint.MiddleCenter
                                                            Dimension1.TextRotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                            Dimension1.Dimasz = 2 'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                            Dimension1.Dimdec = 4 'Sets the number of decimal places displayed for the primary units of a dimension
                                                            Dimension1.Dimtxt = 4 'Specifies the height of dimension text, unless the current text style has a fixed height
                                                            add_extra_param_to_dim(Dimension1, ThisDrawing)
                                                            BTrecord.AppendEntity(Dimension1)
                                                            Trans1.AddNewlyCreatedDBObject(Dimension1, True)
                                                        End If


                                                    End If


                                                Next

                                                If Poly1.Closed = True Then
                                                    Dim Mid_point_poly1 As New Point3d
                                                    Dim Line1_ws As New Line(Poly1.GetPointAtDist(0), Poly1.GetPointAtParameter(Poly1.NumberOfVertices - 1))
                                                    If Line1_ws.Length > 0.01 Then
                                                        Mid_point_poly1 = Line1_ws.GetPointAtDist(Line1_ws.Length / 2)
                                                        Dim Bearing1 As Double = Round(GET_Bearing_rad(Line1_ws.StartPoint.X, Line1_ws.StartPoint.Y, Line1_ws.EndPoint.X, Line1_ws.EndPoint.Y), 4)
                                                        Dim Pt_on_poly As New Point3d
                                                        Pt_on_poly = PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, True)
                                                        Dim Paramcl As Double = PolyCL.GetParameterAtPoint(Pt_on_poly)
                                                        Dim PT_i As New Point3d
                                                        PT_i = PolyCL.GetPointAtParameter(Floor(Paramcl))

                                                        Dim PT_i_1 As New Point3d
                                                        PT_i_1 = PolyCL.GetPointAtParameter(Ceiling(Paramcl))

                                                        If Floor(Paramcl) = Ceiling(Paramcl) Then
                                                            If Floor(Paramcl) >= 1 And Floor(Paramcl) < PolyCL.NumberOfVertices - 1 Then
                                                                PT_i_1 = PolyCL.GetPointAtParameter(Floor(Paramcl) + 1)
                                                            Else
                                                                PT_i = PolyCL.GetPointAtParameter(0)
                                                                PT_i_1 = PolyCL.GetPointAtParameter(1)
                                                            End If
                                                        End If


                                                        Dim Bearing_cl As Double = Round(GET_Bearing_rad(PT_i.X, PT_i.Y, PT_i_1.X, PT_i_1.Y), 4)

                                                        Do Until Bearing_cl < Round(PI, 4)
                                                            If Bearing_cl >= Round(PI, 4) Then
                                                                Bearing_cl = Round(Bearing_cl - Round(PI, 4), 4)
                                                            End If
                                                        Loop

                                                        Do Until Bearing1 < Round(PI, 4)
                                                            If Bearing1 >= Round(PI, 4) Then
                                                                Bearing1 = Round(Bearing1 - Round(PI, 4), 4)
                                                            End If
                                                        Loop


                                                        If Bearing_cl = Bearing1 Or Bearing_cl + Round(PI, 4) = Bearing1 Or Bearing_cl - Round(PI, 4) = Bearing1 Then
                                                            Dim Line1 As New Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, True))
                                                            Line1.Layer = No_plot
                                                            Dim Offset As Double = Line1.Length
                                                            Dim a As Double = Round(Offset, 3)
                                                            Dim b As Double = Round(Offset, 0)
                                                            If Not a = b Then
                                                                Dim Dimension1 As New RotatedDimension
                                                                Dimension1.Layer = No_plot
                                                                Dimension1.XLine1Point = Line1.StartPoint
                                                                Dimension1.XLine2Point = Line1.EndPoint
                                                                Dimension1.Rotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                                Dimension1.DimLinePoint = Line1.StartPoint
                                                                Dimension1.UsingDefaultTextPosition = True
                                                                Dimension1.TextAttachment = AttachmentPoint.MiddleCenter
                                                                Dimension1.TextRotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                                Dimension1.Dimasz = 2 'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                                Dimension1.Dimdec = 4 'Sets the number of decimal places displayed for the primary units of a dimension
                                                                Dimension1.Dimtxt = 4 'Specifies the height of dimension text, unless the current text style has a fixed height
                                                                add_extra_param_to_dim(Dimension1, ThisDrawing)
                                                                BTrecord.AppendEntity(Dimension1)
                                                                Trans1.AddNewlyCreatedDBObject(Dimension1, True)
                                                            End If
                                                        Else


                                                            Dim Line1 As New Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, True))
                                                            Line1.Layer = No_plot
                                                            Dim L1 As Double = Line1.Length

                                                            Dim a As Double = Round((Bearing1 + Round(PI, 4) / 2), 4)
                                                            Dim b As Double = Round(Bearing_cl, 4)
                                                            Dim c As Double = Round((Bearing1 - Round(PI, 4) / 2), 4)
                                                            Dim d As Double = Round(Bearing1, 4)

                                                            Dim Add_L As Boolean = False

                                                            If Not a = b And Not c = b And Not d = b Then
                                                                Add_L = True
                                                            End If
                                                            If Round(L1, 0) = Round(L1, 3) Then
                                                                Add_L = False
                                                            End If

                                                            If Add_L = True Then
                                                                Dim Dimension1 As New RotatedDimension
                                                                Dimension1.Layer = No_plot
                                                                Dimension1.XLine1Point = Line1.StartPoint
                                                                Dimension1.XLine2Point = Line1.EndPoint
                                                                Dimension1.Rotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                                Dimension1.DimLinePoint = Line1.StartPoint
                                                                Dimension1.UsingDefaultTextPosition = True
                                                                Dimension1.TextAttachment = AttachmentPoint.MiddleCenter
                                                                Dimension1.TextRotation = GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y)
                                                                Dimension1.Dimasz = 2 'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                                Dimension1.Dimdec = 4 'Sets the number of decimal places displayed for the primary units of a dimension
                                                                Dimension1.Dimtxt = 4 'Specifies the height of dimension text, unless the current text style has a fixed height
                                                                add_extra_param_to_dim(Dimension1, ThisDrawing)
                                                                BTrecord.AppendEntity(Dimension1)
                                                                Trans1.AddNewlyCreatedDBObject(Dimension1, True)
                                                            End If

                                                        End If


                                                    End If


                                                End If


                                            End If


                                        End If





                                    Next


                                    Trans1.Commit()
                                End If

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
            Freeze_operations = False


        End If
    End Sub



    Private Sub Button_done_Click(sender As Object, e As EventArgs) Handles Button_done.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Dim No_plot As String = "NO PLOT"
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


            Editor1.SetImpliedSelection(Empty_array)
            Try
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()



              

                        Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                        Dim Lista1 As New Specialized.StringCollection
                        For Each Combo1 As Windows.Forms.Control In Panel_layers.Controls
                            If TypeOf Combo1 Is Windows.Forms.ComboBox Then
                                Dim Text1 As String = Combo1.Text
                                If Not Text1 = "" Then
                                    If Lista1.Contains(Text1) = False Then
                                        Lista1.Add(Text1)
                                    End If
                                End If
                            End If
                        Next
                        If Lista1.Count > 0 Then
                            For Each objID As ObjectId In BTrecord
                                Dim Ent1 As Entity = Trans1.GetObject(objID, OpenMode.ForRead)

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.RotatedDimension And Ent1.Layer = No_plot Then
                                    Ent1.UpgradeOpen()
                                    Ent1.Erase()
                                End If

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Region And Ent1.Layer = No_plot Then
                                    Ent1.UpgradeOpen()
                                    Ent1.Erase()
                                End If
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim Poly1 As Polyline = Ent1
                                    If Lista1.Contains(Poly1.Layer) = True Then
                                        Poly1.UpgradeOpen()
                                        Poly1.ColorIndex = 256
                                    End If
                                End If

                            Next
                        End If



                        Trans1.Commit()


                    End Using
                        End Using





                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)

            End Try
            Freeze_operations = False


        End If
    End Sub
End Class