Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class WORKSPACE_FORM

    Private Sub WorkSpace_form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ComboBox_drawing_units.SelectedIndex = 0
        ComboBox_units.SelectedIndex = 0
        Panel_options_for_start_end.Visible = False
        Panel_buffer.Visible = False

    End Sub


    Private Sub Button_draw_WS_Click(sender As System.Object, e As System.EventArgs) Handles Button_draw_WS.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Empty_array = Nothing
        Editor1.SetImpliedSelection(Empty_array)

        Dim CSF As Double = 1
        Dim Nume_layer_NO_PLOT As String = "NO PLOT"
        Dim Nume_layer_WORKSPACE As String = "TempWS"
        Dim factor_conversie As Double = 1
        If ComboBox_units.Text = "Foot" Then
            If ComboBox_drawing_units.Text = "Meter" Then
                factor_conversie = 0.3048
            End If
        End If
        If ComboBox_units.Text = "Meter" Then
            If ComboBox_drawing_units.Text = "Foot" Then
                factor_conversie = 3.28084
            End If
        End If

        If IsNumeric(TextBox_WS_latime.Text) = False Then
            MsgBox("Please specify the width:")
            TextBox_WS_latime.SelectAll()
            Exit Sub
        End If

        Dim Lungime, Latime As Double

        Latime = CDbl(TextBox_WS_latime.Text)
        If Latime <= 0 Then
            MsgBox("Please specify the width:")
            TextBox_WS_latime.SelectAll()
            Exit Sub
        End If

        If CheckBox_start_end.Checked = False Then
            If IsNumeric(TextBox_WS_lungime.Text) = False Then
                MsgBox("Please specify the length:")
                TextBox_WS_lungime.SelectAll()
                Exit Sub
            End If
            Lungime = CDbl(TextBox_WS_lungime.Text)

            If Lungime <= 0 Then
                MsgBox("Please specify the length:")
                TextBox_WS_lungime.SelectAll()
                Exit Sub
            End If
        End If


        If CheckBox_middle_point.Checked = False And CheckBox_Start_point.Checked = False And CheckBox_start_end.Checked = False And CheckBox_buffer.Checked = False Then
            CheckBox_Start_point.Checked = True
        End If


        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock
            Lock1 = ThisDrawing.LockDocument
            Using Lock1


                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select polyline for drafting the workspace:"
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

                    Dim Ent1 As Entity = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)


                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        Dim Poly_base As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent1
                        Dim Poly_offset As New Autodesk.AutoCAD.DatabaseServices.Polyline

                        Creaza_layer(Nume_layer_NO_PLOT, 40, "", False)
                        Creaza_layer(Nume_layer_WORKSPACE, 3, "", True)



                        Latime = Latime * CSF * factor_conversie
                        Lungime = Lungime * CSF * factor_conversie

                        Dim Prompt_pt_point1 As String = vbLf & "Specify start position on the polyline:"
                        If CheckBox_middle_point.Checked = True Then
                            Prompt_pt_point1 = vbLf & "Specify the middle position on the polyline:"
                        End If

                        Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(Prompt_pt_point1)
                        PP1.AllowNone = False

                        Dim Curba1 As Curve
                        Dim Curba2 As Curve

                        Dim Point1_on_base As New Point3d
                        Dim Point2_on_base As New Point3d

                        If CheckBox_select_two_crossing_objects.Checked = False And CheckBox_buffer.Checked = False Then
                            Point1 = Editor1.GetPoint(PP1)
                            If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Exit Sub
                            End If
                            Point1_on_base = Point1.Value
                        End If


                        If CheckBox_start_end.Checked = True Then
                            If CheckBox_select_1crossing_object.Checked = False And CheckBox_select_two_crossing_objects.Checked = False Then
                                Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify END position on the polyline:")
                                PP2.AllowNone = False
                                PP2.BasePoint = Point1_on_base
                                PP2.UseBasePoint = True
                                Point2 = Editor1.GetPoint(PP2)
                                If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Exit Sub
                                End If
                                Point2_on_base = Point2.Value
                            End If




                            If CheckBox_select_1crossing_object.Checked = True Then
                                Dim Crossing_obj_prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbCr & "Select crossing object")
                                Crossing_obj_prompt.SetRejectMessage(vbLf & "Select a line or a polyline")
                                Crossing_obj_prompt.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Curve), False)
                                Dim Crossing_result As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Crossing_obj_prompt)
                                If Crossing_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Exit Sub
                                End If
                                Dim ent_xing As Entity = Trans1.GetObject(Crossing_result.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If TypeOf ent_xing Is Curve Then
                                    Curba1 = ent_xing
                                    Dim Col_int_base As New Point3dCollection
                                    Poly_base.IntersectWith(Curba1, Intersect.OnBothOperands, Col_int_base, IntPtr.Zero, IntPtr.Zero)
                                    If Col_int_base.Count = 0 Then
                                        Exit Sub
                                    End If
                                    Point2_on_base = Col_int_base(0)

                                End If
                            End If


                            If CheckBox_select_two_crossing_objects.Checked = True Then
                                Dim Crossing_obj_prompt1 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbCr & "Select crossing object no 1")
                                Crossing_obj_prompt1.SetRejectMessage(vbLf & "Select a line or a polyline")
                                Crossing_obj_prompt1.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Curve), False)
                                Dim Crossing_result1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Crossing_obj_prompt1)
                                If Crossing_result1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Dim ent_xing1 As Entity = Trans1.GetObject(Crossing_result1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    If TypeOf ent_xing1 Is Curve Then
                                        Curba1 = ent_xing1
                                        Dim Col_int_base1 As New Point3dCollection
                                        Poly_base.IntersectWith(Curba1, Intersect.OnBothOperands, Col_int_base1, IntPtr.Zero, IntPtr.Zero)
                                        If Not Col_int_base1.Count = 0 Then
                                            Point1_on_base = Col_int_base1(0)
                                        End If
                                    End If
                                End If




                                Dim Crossing_obj_prompt2 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbCr & "Select crossing object no 2")
                                Crossing_obj_prompt2.SetRejectMessage(vbLf & "Select a line or a polyline")
                                Crossing_obj_prompt2.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Curve), False)
                                Dim Crossing_result2 As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Crossing_obj_prompt2)
                                If Crossing_result2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Dim ent_xing2 As Entity = Trans1.GetObject(Crossing_result2.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    If TypeOf ent_xing2 Is Curve Then
                                        Curba2 = ent_xing2
                                        Dim Col_int_base2 As New Point3dCollection
                                        Poly_base.IntersectWith(Curba2, Intersect.OnBothOperands, Col_int_base2, IntPtr.Zero, IntPtr.Zero)
                                        If Not Col_int_base2.Count = 0 Then
                                            Point2_on_base = Col_int_base2(0)
                                        End If
                                    End If
                                End If


                            End If




                        End If


                        Dim Pointdir As Autodesk.AutoCAD.EditorInput.PromptPointResult

                        If CheckBox_Start_point.Checked = True Then
                            Dim PPdir As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please specify direction!:")
                            PPdir.BasePoint = Point1_on_base
                            PPdir.UseBasePoint = True
                            PPdir.AllowNone = False

                            Pointdir = Editor1.GetPoint(PPdir)
                            If Not Pointdir.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Exit Sub
                            End If
                        End If ' asta e If CheckBox_Start_point.Checked = True

                        Dim PointSide As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim PromtPointSide As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify workspace side")

                        PromtPointSide.AllowNone = False
                        If CheckBox_buffer.Checked = False Then
                            PromtPointSide.BasePoint = Point1_on_base
                            PromtPointSide.UseBasePoint = True
                            PointSide = Editor1.GetPoint(PromtPointSide)
                            If Not PointSide.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Exit Sub
                            End If
                        End If


                        Dim Xside, Yside As Double
                        If CheckBox_buffer.Checked = False Then
                            Xside = PointSide.Value.X
                            Yside = PointSide.Value.Y
                        End If

                        Dim Xside_pe_poly, Yside_pe_poly As Double
                        If CheckBox_buffer.Checked = False Then
                            Xside_pe_poly = Poly_base.GetClosestPointTo(PointSide.Value, Vector3d.ZAxis, False).X
                            Yside_pe_poly = Poly_base.GetClosestPointTo(PointSide.Value, Vector3d.ZAxis, False).Y

                            If ((Xside - Xside_pe_poly) ^ 2 + (Yside - Yside_pe_poly) ^ 2) ^ 0.5 <= 0.001 Then
                                MsgBox("You picked too close")
                                Exit Sub
                            End If

                            Dim Line_to_be_scaled As New Line(New Point3d(Xside_pe_poly, Yside_pe_poly, 0), New Point3d(Xside, Yside, 0))
                            Line_to_be_scaled.TransformBy(Matrix3d.Scaling(1000, New Point3d(Xside_pe_poly, Yside_pe_poly, 0)))
                            Xside = Line_to_be_scaled.EndPoint.X
                            Yside = Line_to_be_scaled.EndPoint.Y


                        End If


                        Dim X0, Y0 As Double
                        Dim X1, Y1 As Double



                        Dim Lungime_de_la_Start As Double
                        Dim Lungime_de_la_End As Double
                        Dim Lungime_de_la_Directie As Double
                        Dim Lungime_poly_base As Double = Poly_base.Length






                        Dim X_dir, Y_dir As Double


                        If CheckBox_buffer.Checked = True Then
                            Dim Buffer_obj_prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbCr & "Select the feature for applying the buffer")
                            Buffer_obj_prompt.SetRejectMessage(vbLf & "Select a line or a polyline")
                            Buffer_obj_prompt.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Curve), False)
                            Dim buffer_result As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(Buffer_obj_prompt)
                            If buffer_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Exit Sub
                            End If
                            Dim ent_Buffer As Entity = Trans1.GetObject(buffer_result.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            Dim CenterLine_obj_prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbCr & "Select CenterLine")
                            CenterLine_obj_prompt.SetRejectMessage(vbLf & "Select a polyline")
                            CenterLine_obj_prompt.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), False)
                            Dim CenterLine_result As Autodesk.AutoCAD.EditorInput.PromptEntityResult = Editor1.GetEntity(CenterLine_obj_prompt)
                            If CenterLine_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Exit Sub
                            End If
                            Dim ent_CenterLine As Entity = Trans1.GetObject(CenterLine_result.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            If TypeOf ent_Buffer Is Curve And TypeOf ent_CenterLine Is Curve Then
                                Curba1 = ent_Buffer
                                Dim Col_int_base As New Point3dCollection
                                Poly_base.IntersectWith(Curba1, Intersect.OnBothOperands, Col_int_base, IntPtr.Zero, IntPtr.Zero)
                                If Col_int_base.Count = 0 Then
                                    Exit Sub
                                End If

                                Dim PPdir As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please specify workspace direction!:")

                                PPdir.BasePoint = Col_int_base(0)
                                PPdir.UseBasePoint = True

                                PPdir.AllowNone = False
                                Pointdir = Editor1.GetPoint(PPdir)
                                If Not Pointdir.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Exit Sub
                                End If

                                Dim Point_directie_pe_baza As New Point3d
                                Point_directie_pe_baza = Poly_base.GetClosestPointTo(Pointdir.Value, Vector3d.ZAxis, False)
                                Dim Len_DIR As Double = 100000
                                Dim pct_index_aproape As Integer
                                If Col_int_base.Count > 0 Then
                                    For k = 0 To Col_int_base.Count - 1
                                        'MsgBox(Point_directie_pe_baza.GetVectorTo(Col_int_offset(k)).Length)
                                        If Point_directie_pe_baza.GetVectorTo(Col_int_base(k)).Length < Len_DIR Then
                                            Len_DIR = Point_directie_pe_baza.GetVectorTo(Col_int_base(k)).Length
                                            pct_index_aproape = k
                                        End If
                                    Next
                                End If

                                Dim Param_INT As Double = Poly_base.GetParameterAtPoint(Col_int_base(pct_index_aproape))
                                Dim Param_pct_dir As Double = Poly_base.GetParameterAtPoint(Point_directie_pe_baza)


                                If Param_pct_dir < Param_INT Then
                                    Point_directie_pe_baza = Poly_base.GetPointAtParameter(0)
                                Else
                                    Point_directie_pe_baza = Poly_base.GetPointAtDist(Poly_base.Length)
                                End If


                                PromtPointSide.BasePoint = Col_int_base(0)
                                PromtPointSide.UseBasePoint = True
                                PointSide = Editor1.GetPoint(PromtPointSide)
                                If Not PointSide.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Exit Sub
                                End If


                                Dim Object_colection2 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = Curba1.GetOffsetCurves(CDbl(TextBox_buffer.Text) * Directie_offset(Curba1.ObjectId, Pointdir.Value))

                                Dim Poly_offsetata As New Polyline
                                Poly_offsetata = Object_colection2(0)
                                Poly_offsetata.ColorIndex = 1
                                BTrecord.AppendEntity(Poly_offsetata)
                                Trans1.AddNewlyCreatedDBObject(Poly_offsetata, True)

                                Dim CL As Curve = ent_CenterLine
                                Dim Param_min As Double = 1000000000000
                                Dim Param_max As Double = -1
                                Dim Nod_min As Double
                                Dim Nod_max As Double

                                For k = 0 To Poly_offsetata.NumberOfVertices - 1
                                    Dim Nod_poly As New Point3d
                                    Nod_poly = Poly_offsetata.GetPoint3dAt(k)
                                    Dim Point_on_CL As New Point3d
                                    Point_on_CL = CL.GetClosestPointTo(Nod_poly, Vector3d.ZAxis, False)
                                    If CL.GetParameterAtPoint(Point_on_CL) < Param_min Then
                                        Param_min = CL.GetParameterAtPoint(Point_on_CL)
                                        Nod_min = k
                                    End If
                                    If CL.GetParameterAtPoint(Point_on_CL) > Param_max Then
                                        Param_max = CL.GetParameterAtPoint(Point_on_CL)
                                        Nod_max = k
                                    End If
                                Next

                                Dim Linie_min As New Line(CL.GetPointAtParameter(Param_min), Poly_offsetata.GetPointAtParameter(Nod_min))
                                Dim Linie_max As New Line(CL.GetPointAtParameter(Param_max), Poly_offsetata.GetPointAtParameter(Nod_max))

                                'BTrecord.AppendEntity(Linie_min)
                                'Trans1.AddNewlyCreatedDBObject(Linie_min, True)
                                'BTrecord.AppendEntity(Linie_max)
                                'Trans1.AddNewlyCreatedDBObject(Linie_max, True)


                                Dim Int_min As New Point3dCollection
                                Dim Int_max As New Point3dCollection

                                Poly_base.IntersectWith(Linie_min, Intersect.ExtendArgument, Int_min, IntPtr.Zero, IntPtr.Zero)
                                Poly_base.IntersectWith(Linie_max, Intersect.ExtendArgument, Int_max, IntPtr.Zero, IntPtr.Zero)

                                ' MsgBox(Int_min.Count & " min" & vbCrLf & Int_max.Count & " max")

                                Dim P_min As Double = 0
                                Dim P_max As Double = Poly_base.GetParameterAtDistance(Poly_base.Length)

                                If Int_min.Count > 0 And Int_max.Count > 0 Then
                                    P_min = Poly_base.GetParameterAtPoint(Int_min(0))
                                    P_max = Poly_base.GetParameterAtPoint(Int_max(0))
                                    If P_min > P_max Then
                                        Dim T As Double = P_min
                                        P_min = P_max
                                        P_max = T
                                    End If

                                Else
                                    Trans1.Commit()
                                    Exit Sub

                                End If

                                If Param_pct_dir < Param_INT Then
                                    Point1_on_base = Poly_base.GetPointAtParameter(P_min)
                                Else
                                    Point1_on_base = Poly_base.GetPointAtParameter(P_max)
                                End If







                                X0 = Point1_on_base.X
                                Y0 = Point1_on_base.Y
                                Lungime_de_la_Start = Poly_base.GetDistAtPoint(New Point3d(X0, Y0, 0))

                                X_dir = Point_directie_pe_baza.X
                                Y_dir = Point_directie_pe_baza.Y
                                Lungime_de_la_Directie = Poly_base.GetDistAtPoint(New Point3d(X_dir, Y_dir, 0))

                                If Lungime_de_la_Directie > Lungime_de_la_Start Then
                                    If (Lungime_poly_base - Lungime_de_la_Start) > Lungime Then
                                        X1 = Poly_base.GetPointAtDist(Lungime_de_la_Start + Lungime).X
                                        Y1 = Poly_base.GetPointAtDist(Lungime_de_la_Start + Lungime).Y
                                    Else
                                        X1 = Poly_base.EndPoint.X
                                        Y1 = Poly_base.EndPoint.Y
                                        Lungime = Lungime_poly_base - Lungime_de_la_Start
                                    End If
                                Else
                                    X1 = X0
                                    Y1 = Y0

                                    If Lungime_de_la_Start > Lungime Then
                                        X0 = Poly_base.GetPointAtDist(Lungime_de_la_Start - Lungime).X
                                        Y0 = Poly_base.GetPointAtDist(Lungime_de_la_Start - Lungime).Y
                                    Else
                                        X0 = Poly_base.StartPoint.X
                                        Y0 = Poly_base.StartPoint.Y
                                        Lungime = Lungime_de_la_Start
                                    End If
                                    Lungime_de_la_Start = Poly_base.GetDistAtPoint(New Point3d(X0, Y0, 0))

                                End If
                                Lungime_de_la_End = Poly_base.GetDistAtPoint(New Point3d(X1, Y1, 0))

                            End If

                        End If



                        If CheckBox_Start_point.Checked = True Then
                            X0 = Poly_base.GetClosestPointTo(New Point3d(Point1_on_base.X, Point1_on_base.Y, 0), Vector3d.ZAxis, False).X
                            Y0 = Poly_base.GetClosestPointTo(New Point3d(Point1_on_base.X, Point1_on_base.Y, 0), Vector3d.ZAxis, False).Y
                            Lungime_de_la_Start = Poly_base.GetDistAtPoint(New Point3d(X0, Y0, 0))

                            X_dir = Poly_base.GetClosestPointTo(New Point3d(Pointdir.Value.X, Pointdir.Value.Y, 0), Vector3d.ZAxis, False).X
                            Y_dir = Poly_base.GetClosestPointTo(New Point3d(Pointdir.Value.X, Pointdir.Value.Y, 0), Vector3d.ZAxis, False).Y
                            Lungime_de_la_Directie = Poly_base.GetDistAtPoint(New Point3d(X_dir, Y_dir, 0))

                            If Lungime_de_la_Directie > Lungime_de_la_Start Then
                                If (Lungime_poly_base - Lungime_de_la_Start) > Lungime Then
                                    X1 = Poly_base.GetPointAtDist(Lungime_de_la_Start + Lungime).X
                                    Y1 = Poly_base.GetPointAtDist(Lungime_de_la_Start + Lungime).Y
                                Else
                                    X1 = Poly_base.EndPoint.X
                                    Y1 = Poly_base.EndPoint.Y
                                    Lungime = Lungime_poly_base - Lungime_de_la_Start
                                End If
                            Else
                                X1 = X0
                                Y1 = Y0

                                If Lungime_de_la_Start > Lungime Then
                                    X0 = Poly_base.GetPointAtDist(Lungime_de_la_Start - Lungime).X
                                    Y0 = Poly_base.GetPointAtDist(Lungime_de_la_Start - Lungime).Y
                                Else
                                    X0 = Poly_base.StartPoint.X
                                    Y0 = Poly_base.StartPoint.Y
                                    Lungime = Lungime_de_la_Start
                                End If
                                Lungime_de_la_Start = Poly_base.GetDistAtPoint(New Point3d(X0, Y0, 0))

                            End If
                            Lungime_de_la_End = Poly_base.GetDistAtPoint(New Point3d(X1, Y1, 0))

                        End If

                        


                        If CheckBox_middle_point.Checked = True Then
                            X0 = Poly_base.GetClosestPointTo(New Point3d(Point1_on_base.X, Point1_on_base.Y, 0), Vector3d.ZAxis, False).X
                            Y0 = Poly_base.GetClosestPointTo(New Point3d(Point1_on_base.X, Point1_on_base.Y, 0), Vector3d.ZAxis, False).Y
                            Lungime_de_la_Start = Poly_base.GetDistAtPoint(New Point3d(X0, Y0, 0))

                            If Lungime_de_la_Start >= Lungime / 2 Then
                                Lungime_de_la_Start = Lungime_de_la_Start - Lungime / 2
                            Else
                                Exit Sub
                            End If

                            X0 = Poly_base.GetPointAtDist(Lungime_de_la_Start).X
                            Y0 = Poly_base.GetPointAtDist(Lungime_de_la_Start).Y
                            If (Lungime_poly_base - Lungime_de_la_Start) > Lungime Then
                                X1 = Poly_base.GetPointAtDist(Lungime_de_la_Start + Lungime).X
                                Y1 = Poly_base.GetPointAtDist(Lungime_de_la_Start + Lungime).Y
                            Else
                                X1 = Poly_base.EndPoint.X
                                Y1 = Poly_base.EndPoint.Y
                            End If

                            Lungime_de_la_End = Poly_base.GetDistAtPoint(New Point3d(X1, Y1, 0))
                        End If

                        Dim Reverse_start_end As Boolean = False


                        'aici
                        If CheckBox_start_end.Checked = True Then
                            X0 = Poly_base.GetClosestPointTo(New Point3d(Point1_on_base.X, Point1_on_base.Y, 0), Vector3d.ZAxis, False).X
                            Y0 = Poly_base.GetClosestPointTo(New Point3d(Point1_on_base.X, Point1_on_base.Y, 0), Vector3d.ZAxis, False).Y
                            Lungime_de_la_Start = Poly_base.GetDistAtPoint(New Point3d(X0, Y0, 0))

                            X1 = Poly_base.GetClosestPointTo(Point2_on_base, Vector3d.ZAxis, False).X
                            Y1 = Poly_base.GetClosestPointTo(Point2_on_base, Vector3d.ZAxis, False).Y
                            Lungime_de_la_End = Poly_base.GetDistAtPoint(New Point3d(X1, Y1, 0))
                            If Lungime_de_la_End < Lungime_de_la_Start Then
                                Dim Xt, Yt As Double
                                Xt = X0
                                Yt = Y0
                                X0 = X1
                                Y0 = Y1
                                X1 = Xt
                                Y1 = Yt
                                Xt = Lungime_de_la_Start
                                Lungime_de_la_Start = Lungime_de_la_End
                                Lungime_de_la_End = Xt
                                Reverse_start_end = True
                            End If
                            Lungime = Lungime_de_la_End - Lungime_de_la_Start
                        End If


                        Dim indexstart As Double = Poly_base.GetParameterAtDistance(Lungime_de_la_Start)
                        Dim indexend As Double = Poly_base.GetParameterAtDistance(Lungime_de_la_End)


                        Dim Colectie_puncte_on_base As New Point2dCollection

                        If Floor(indexend) - Floor(indexstart) = 0 Then
                            Colectie_puncte_on_base.Add(New Point2d(X0, Y0))
                            Colectie_puncte_on_base.Add(New Point2d(X1, Y1))
                        Else
                            Dim Start1, End1 As Integer
                            If Abs(indexstart - Round(indexstart, 0)) < 0.01 Then
                                Start1 = CInt(Round(indexstart, 0)) + 1
                            Else
                                Start1 = Floor(indexstart) + 1
                            End If
                            If Abs(indexend - Round(indexend, 0)) < 0.01 Then
                                End1 = CInt(Round(indexend, 0)) - 1
                            Else
                                End1 = Ceiling(indexend) - 1
                            End If
                            Colectie_puncte_on_base.Add(New Point2d(X0, Y0))

                            For i = Start1 To End1
                                Colectie_puncte_on_base.Add(Poly_base.GetPoint2dAt(i))


                            Next
                            Colectie_puncte_on_base.Add(New Point2d(X1, Y1))

                        End If




                        Dim Polilinie_break As New Polyline

                        If Colectie_puncte_on_base.Count > 0 Then
                            For i = 0 To Colectie_puncte_on_base.Count - 1
                                Polilinie_break.AddVertexAt(i, Colectie_puncte_on_base(i), 0, 0, 0)

                            Next
                        End If

                        If Reverse_start_end = True Then
                            Polilinie_break.ReverseCurve()
                        End If




                        Dim Object_colection1 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection

                        Object_colection1 = Polilinie_break.GetOffsetCurves(Latime * Directie_offset(Poly_base.ObjectId, PointSide.Value))
                        Poly_offset = Object_colection1(0)


                        Dim Poly_ws As New Polyline
                        Dim Index_ws As Integer = 0
                        For i = 0 To Poly_offset.NumberOfVertices - 1
                            Poly_ws.AddVertexAt(Index_ws, Poly_offset.GetPoint2dAt(i), 0, 0, 0)
                            Index_ws = Index_ws + 1
                        Next
                        For i = Polilinie_break.NumberOfVertices - 1 To 0 Step -1
                            Poly_ws.AddVertexAt(Index_ws, Polilinie_break.GetPoint2dAt(i), 0, 0, 0)
                            Index_ws = Index_ws + 1
                        Next

                        If CheckBox_select_1crossing_object.Checked = True Then
                            If IsNothing(Curba1) = False Then
                                Dim Linie_sus As New Line(Poly_ws.GetPoint3dAt(Poly_offset.NumberOfVertices - 2), Poly_ws.GetPoint3dAt(Poly_offset.NumberOfVertices - 1))
                                Dim Col_int As New Point3dCollection
                                Linie_sus.IntersectWith(Curba1, Intersect.ExtendBoth, Col_int, IntPtr.Zero, IntPtr.Zero)
                                If Col_int.Count > 0 Then
                                    Poly_ws.RemoveVertexAt(Poly_offset.NumberOfVertices - 1)
                                    Poly_ws.AddVertexAt(Poly_offset.NumberOfVertices - 1, New Point2d(Col_int(0).X, Col_int(0).Y), 0, 0, 0)
                                End If
                            End If
                        End If


                        If CheckBox_select_two_crossing_objects.Checked = True Then
                            If IsNothing(Curba1) = False Then
                                Dim Linie_sus As New Line(Poly_ws.GetPoint3dAt(0), Poly_ws.GetPoint3dAt(1))
                                Dim Col_int As New Point3dCollection
                                Linie_sus.IntersectWith(Curba1, Intersect.ExtendBoth, Col_int, IntPtr.Zero, IntPtr.Zero)
                                If Col_int.Count > 0 Then
                                    Poly_ws.RemoveVertexAt(0)
                                    Poly_ws.AddVertexAt(0, New Point2d(Col_int(0).X, Col_int(0).Y), 0, 0, 0)
                                End If
                            End If

                            If IsNothing(Curba2) = False Then
                                Dim Linie_sus As New Line(Poly_ws.GetPoint3dAt(Poly_offset.NumberOfVertices - 2), Poly_ws.GetPoint3dAt(Poly_offset.NumberOfVertices - 1))
                                Dim Col_int As New Point3dCollection
                                Linie_sus.IntersectWith(Curba2, Intersect.ExtendBoth, Col_int, IntPtr.Zero, IntPtr.Zero)
                                If Col_int.Count > 0 Then
                                    Poly_ws.RemoveVertexAt(Poly_offset.NumberOfVertices - 1)
                                    Poly_ws.AddVertexAt(Poly_offset.NumberOfVertices - 1, New Point2d(Col_int(0).X, Col_int(0).Y), 0, 0, 0)
                                End If
                            End If
                        End If

                        Poly_ws.Closed = True
                        Poly_ws.Layer = Nume_layer_WORKSPACE

                        BTrecord.AppendEntity(Poly_ws)
                        Trans1.AddNewlyCreatedDBObject(Poly_ws, True)

                        If CheckBox_select_1crossing_object.Checked = True Then
                            If CheckBox_measure_on_middle.Checked = True Then
                                Dim Object_colection2 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = Polilinie_break.GetOffsetCurves(0.5 * Latime * Directie_offset(Poly_base.ObjectId, PointSide.Value))

                                Poly_offset = Object_colection2(0)
                                Dim Intcol2 As New Point3dCollection
                                Poly_offset.IntersectWith(Curba1, Intersect.ExtendThis, Intcol2, IntPtr.Zero, IntPtr.Zero)
                                If Intcol2.Count > 0 Then
                                    Dim nr_vert As Integer = Poly_offset.NumberOfVertices

                                    Poly_offset.RemoveVertexAt(nr_vert - 1)
                                    Poly_offset.AddVertexAt(nr_vert - 1, New Point2d(Intcol2(0).X, Intcol2(0).Y), 0, 0, 0)


                                    Lungime = Poly_offset.Length
                                End If


                                '*sterge
                                'Poly_offset.Layer = Nume_layer_NO_PLOT
                                'BTrecord.AppendEntity(Poly_offset)
                                'Trans1.AddNewlyCreatedDBObject(Poly_offset, True)

                            End If
                        End If

                        If CheckBox_select_two_crossing_objects.Checked = True Then
                            If CheckBox_measure_on_middle.Checked = True Then

                                Dim Object_colection2 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = Polilinie_break.GetOffsetCurves(0.5 * Latime * Directie_offset(Poly_base.ObjectId, PointSide.Value))

                                Poly_offset = Object_colection2(0)
                                Dim Intcol1 As New Point3dCollection
                                Dim Intcol2 As New Point3dCollection
                                Poly_offset.IntersectWith(Curba1, Intersect.ExtendThis, Intcol1, IntPtr.Zero, IntPtr.Zero)
                                Poly_offset.IntersectWith(Curba2, Intersect.ExtendThis, Intcol2, IntPtr.Zero, IntPtr.Zero)
                                If Intcol1.Count > 0 And Intcol2.Count > 0 Then
                                    Dim nr_vert As Integer = Poly_offset.NumberOfVertices

                                    Poly_offset.RemoveVertexAt(0)
                                    Poly_offset.AddVertexAt(0, New Point2d(Intcol1(0).X, Intcol1(0).Y), 0, 0, 0)

                                    Poly_offset.RemoveVertexAt(nr_vert - 1)
                                    Poly_offset.AddVertexAt(nr_vert - 1, New Point2d(Intcol2(0).X, Intcol2(0).Y), 0, 0, 0)

                                    Lungime = Poly_offset.Length
                                End If




                            End If
                        End If

                        Dim Nr_laturi As Integer
                        Dim Xmtext, Ymtext As Double
                        Nr_laturi = Poly_ws.NumberOfVertices
                        Dim Nr_temp_x, Nr_temp_y As Integer

                        For i = 0 To Nr_laturi - 1
                            Nr_temp_x = Nr_temp_x + Poly_ws.GetPoint2dAt(i).X
                            Nr_temp_y = Nr_temp_y + Poly_ws.GetPoint2dAt(i).Y
                        Next
                        Xmtext = Nr_temp_x / Nr_laturi
                        Ymtext = Nr_temp_y / Nr_laturi

                        Dim Mtext1 As New Autodesk.AutoCAD.DatabaseServices.MText
                        Mtext1.Layer = Nume_layer_NO_PLOT
                        Mtext1.Location = New Point3d(Xmtext, Ymtext, 0)
                        Mtext1.Contents = Round(Latime / CSF, 1) & " x " & Round(Lungime / CSF, 1)
                        If ComboBox_drawing_units.Text = "Foot" Then
                            Mtext1.Contents = Round(Latime / CSF, 1) & "' x " & Round(Lungime / CSF, 1) & "'"
                        End If
                        Mtext1.TextHeight = 2.5
                        Mtext1.Attachment = Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.MiddleCenter
                        BTrecord.AppendEntity(Mtext1)
                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                        Trans1.Commit()

                        Editor1.SetImpliedSelection(Empty_array)

                    End If ' asta e   de la If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline 
                End Using ' asta e   de la trans1
            End Using ' asta e   de la lock
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
            Editor1.WriteMessage(vbLf & "Command:")

        End Try

    End Sub


    Private Sub WorkSpace_form_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        If ComboBox_drawing_units.SelectedIndex = 1 Then
            ComboBox_drawing_units.SelectedIndex = 0
            ComboBox_units.SelectedIndex = 0
        Else
            ComboBox_drawing_units.SelectedIndex = 1
            ComboBox_units.SelectedIndex = 1
        End If
    End Sub



    Private Sub TextBox_WS_latime_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_WS_latime.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            TextBox_WS_lungime.SelectAll()
            TextBox_WS_lungime.Focus()
        End If
    End Sub

    Private Sub TextBox_WS_lungime_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_WS_lungime.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            Button_draw_WS_Click(sender, e)
        End If
    End Sub


    Private Sub Panel_Click(sender As Object, e As EventArgs) Handles Panel3.Click
        If ComboBox_drawing_units.SelectedIndex = 1 Then
            ComboBox_drawing_units.SelectedIndex = 0
            ComboBox_units.SelectedIndex = 0
        Else
            ComboBox_drawing_units.SelectedIndex = 1
            ComboBox_units.SelectedIndex = 1
        End If
    End Sub


    Private Sub CheckBox_start_end_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_start_end.CheckedChanged
        If CheckBox_start_end.Checked = True Then
            Panel_options_for_start_end.Visible = True
            CheckBox_middle_point.Checked = False
            CheckBox_Start_point.Checked = False
            CheckBox_measure_Bottom.Checked = True
            CheckBox_select_1crossing_object.Checked = False
            CheckBox_select_two_crossing_objects.Checked = False
            CheckBox_buffer.Checked = False
        Else
            Panel_options_for_start_end.Visible = False
            CheckBox_select_1crossing_object.Checked = False
            CheckBox_select_two_crossing_objects.Checked = False
            CheckBox_measure_on_middle.Checked = False
            CheckBox_measure_Bottom.Checked = True
        End If
    End Sub

    Private Sub CheckBox_measure_Bottom_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_measure_Bottom.CheckedChanged
        If CheckBox_measure_Bottom.Checked = True Then
            CheckBox_measure_on_middle.Checked = False
            CheckBox_measure_on_middle.Checked = False

        Else
            CheckBox_measure_on_middle.Checked = True
        End If
    End Sub

    Private Sub CheckBox_measure_on_middle_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_measure_on_middle.CheckedChanged
        If CheckBox_measure_on_middle.Checked = True Then
            If CheckBox_measure_Bottom.Checked = True Then CheckBox_measure_Bottom.Checked = False
        Else
            If CheckBox_measure_Bottom.Checked = False Then CheckBox_measure_Bottom.Checked = True
        End If
    End Sub
    Private Sub CheckBox_select_two_crossing_objects_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_select_two_crossing_objects.CheckedChanged
        If CheckBox_select_two_crossing_objects.Checked = True Then
            CheckBox_select_1crossing_object.Checked = False
        End If

    End Sub
    Private Sub CheckBox_select_1crossing_object_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_select_1crossing_object.CheckedChanged
        If CheckBox_select_1crossing_object.Checked = True Then
            CheckBox_select_two_crossing_objects.Checked = False
        End If

    End Sub
    Private Sub CheckBox_Start_point_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_Start_point.CheckedChanged
        If CheckBox_Start_point.Checked = True Then
            CheckBox_measure_Bottom.Checked = True
            CheckBox_measure_on_middle.Checked = False
            CheckBox_measure_Bottom.Checked = False
            Panel_options_for_start_end.Visible = False
            CheckBox_middle_point.Checked = False
            CheckBox_start_end.Checked = False
            CheckBox_buffer.Checked = False

        End If
    End Sub
    Private Sub CheckBox_middle_point_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_middle_point.CheckedChanged
        If CheckBox_middle_point.Checked = True Then
            CheckBox_measure_Bottom.Checked = True
            CheckBox_measure_on_middle.Checked = False
            CheckBox_measure_Bottom.Checked = False
            Panel_options_for_start_end.Visible = False
            CheckBox_Start_point.Checked = False
            CheckBox_start_end.Checked = False
            CheckBox_buffer.Checked = False
        End If
    End Sub

    Private Sub CheckBox_buffer_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_buffer.CheckedChanged
        If CheckBox_buffer.Checked = True Then
            Panel_buffer.Visible = True
            CheckBox_Start_point.Checked = False
            CheckBox_start_end.Checked = False
        Else
            Panel_buffer.Visible = False
        End If
    End Sub
End Class