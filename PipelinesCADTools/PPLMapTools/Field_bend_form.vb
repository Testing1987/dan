
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Field_bend_form
    Private Sub Field_bend_form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ComboBox_nps.Items.Add("NPS 1/2")
        ComboBox_nps.Items.Add("NPS 1")
        ComboBox_nps.Items.Add("NPS 1.5")
        ComboBox_nps.Items.Add("NPS 2")
        ComboBox_nps.Items.Add("NPS 3")
        ComboBox_nps.Items.Add("NPS 4")
        ComboBox_nps.Items.Add("NPS 6")
        ComboBox_nps.Items.Add("NPS 8")
        ComboBox_nps.Items.Add("NPS 10")
        ComboBox_nps.Items.Add("NPS 12")
        ComboBox_nps.Items.Add("NPS 16")
        ComboBox_nps.Items.Add("NPS 20")
        ComboBox_nps.Items.Add("NPS 24")
        ComboBox_nps.Items.Add("NPS 30")
        ComboBox_nps.Items.Add("NPS 36")
        ComboBox_nps.Items.Add("NPS 42")
        ComboBox_nps.Items.Add("NPS 48")
        ComboBox_nps.SelectedIndex = 13
        TextBox_radius1.Text = 10 * 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(30) / 1000
        TextBox_radius2.Text = 10 * 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(30) / 1000
        Incarca_existing_layers_to_combobox(ComboBox_layer_CL)
        Incarca_existing_layers_to_combobox(ComboBox_layer_OD)
        If ComboBox_layer_CL.Items.Contains("PCENTRE") = True Then
            ComboBox_layer_CL.SelectedIndex = ComboBox_layer_CL.Items.IndexOf("PCENTRE")
        End If
        If ComboBox_layer_OD.Items.Contains("PNEW") = True Then
            ComboBox_layer_OD.SelectedIndex = ComboBox_layer_OD.Items.IndexOf("PNEW")
        End If
    End Sub

    Private Sub Button_draw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_draw.Click
        Try

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument
                Dim R1, R2 As Double
                If IsNumeric(TextBox_radius1.Text) = True And IsNumeric(TextBox_radius2.Text) = True Then
                    R1 = CDbl(TextBox_radius1.Text)
                    R2 = CDbl(TextBox_radius2.Text)
                Else
                    MsgBox("Verify the radius")
                    Exit Sub
                End If

                Dim Lungime_arc1, Lungime_arc2, Lungime_linie, Unghi_arc1, Unghi_arc2 As Double
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


                Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt1.MessageForAdding = vbLf & "Select the Horizontal segment:"

                Object_Prompt1.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt1)




                If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Exit Sub
                End If


                Dim Rezultat11 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt11 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt11.MessageForAdding = vbLf & "Select the second segment:"
                Object_Prompt11.SingleOnly = True
                Rezultat11 = Editor1.GetSelection(Object_Prompt11)




                If Rezultat11.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Exit Sub
                End If

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And IsNothing(Rezultat1) = False And Rezultat11.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And IsNothing(Rezultat11) = False Then


                    Dim Line1, Line2 As Autodesk.AutoCAD.DatabaseServices.Line

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat11.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat1.Value.Item(0)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly1 As Polyline = Ent1
                            Dim Linie1 As New Line(Poly1.GetPoint3dAt(Poly1.NumberOfVertices - 2), Poly1.GetPoint3dAt(Poly1.NumberOfVertices - 1))
                            Ent1 = Linie1
                        End If

                        If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly2 As Polyline = Ent2
                            Dim Linie2 As New Line(Poly2.GetPoint3dAt(Poly2.NumberOfVertices - 2), Poly2.GetPoint3dAt(Poly2.NumberOfVertices - 1))
                            Ent2 = Linie2
                        End If

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                            Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord

                            BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)

                            Line1 = Ent1
                            Line2 = Ent2



                            Dim x01, x02, x03, x04, y01, y02, y03, y04, X5, X6, Y5, Y6, X7, Y7, X8, Y8, X9, Y9 As Double
                            x01 = Line1.StartPoint.X
                            x02 = Line1.EndPoint.X
                            x03 = Line2.StartPoint.X
                            x04 = Line2.EndPoint.X
                            y01 = Line1.StartPoint.Y
                            y02 = Line1.EndPoint.Y
                            y03 = Line2.StartPoint.Y
                            y04 = Line2.EndPoint.Y

                            Dim D1, D2, D3, D4, xt, yt As Double

                            D1 = GET_distanta_Double_XY(x01, y01, x03, y03)
                            D2 = GET_distanta_Double_XY(x01, y01, x04, y04)
                            D3 = GET_distanta_Double_XY(x02, y02, x03, y03)
                            D4 = GET_distanta_Double_XY(x02, y02, x04, y04)

                            If D1 <= D2 Then
                                If D1 <= D3 Then
                                    If D1 <= D4 Then
                                        xt = x01
                                        yt = y01
                                        x01 = x02
                                        y01 = y02
                                        x02 = xt
                                        y02 = yt
                                        'd1=min
                                    Else 'D1 > D4
                                        xt = x03
                                        yt = y03
                                        x03 = x04
                                        y03 = y04
                                        x04 = xt
                                        y04 = yt
                                        'd4=min
                                    End If 'D1 <= D4
                                Else 'D1 > D3
                                    If D3 <= D4 Then
                                        'd3=min
                                    Else 'D3 <= D4
                                        xt = x03
                                        yt = y03
                                        x03 = x04
                                        y03 = y04
                                        x04 = xt
                                        y04 = yt
                                        'D4=min
                                    End If 'D3 <= D4
                                End If 'D1 <= D3

                            Else ' D1 > D2
                                If D2 <= D3 Then
                                    If D2 <= D4 Then
                                        xt = x04
                                        yt = y04
                                        x04 = x03
                                        y04 = y03
                                        x03 = xt
                                        y03 = yt

                                        xt = x01
                                        yt = y01
                                        x01 = x02
                                        y01 = y02
                                        x02 = xt
                                        y02 = yt

                                        'd2=min
                                    Else 'd2 > D4
                                        xt = x03
                                        yt = y03
                                        x03 = x04
                                        y03 = y04
                                        x04 = xt
                                        y04 = yt
                                        'd4=min
                                    End If 'd2 <= D4
                                Else 'd2> D3
                                    If D3 <= D4 Then
                                        'd3=min
                                    Else 'D3 > D4
                                        xt = x03
                                        yt = y03
                                        x03 = x04
                                        y03 = y04
                                        x04 = xt
                                        y04 = yt
                                        'D4=min
                                    End If 'D3 <= D4
                                End If 'd2 <= D3
                            End If



                            Dim x02_extended, y02_extended As Double
                            Dim Lungime_pana_la_3 As Double
                            Dim Dist_01_03 As Double
                            Dist_01_03 = GET_distanta_Double_XY(x01, y01, x03, y03)
                            Dim BearingT1, BearingT2 As Double
                            BearingT1 = GET_Bearing_rad(x01, y01, x03, y03)
                            BearingT2 = GET_Bearing_rad(x01, y01, x02, y02)
                            Dim UnghiT As Double = Abs(BearingT1 - BearingT2)
                            If UnghiT > PI Then UnghiT = 2 * PI - UnghiT
                            Lungime_pana_la_3 = Dist_01_03 * Cos(UnghiT)


                            y02_extended = y01 + Lungime_pana_la_3 * Sin(GET_Bearing_rad(x01, y01, x02, y02))
                            x02_extended = x01 + Lungime_pana_la_3 * Cos(GET_Bearing_rad(x01, y01, x02, y02))



                            Dim Line1_extended As New Autodesk.AutoCAD.DatabaseServices.Line


                            Line1_extended.StartPoint = New Point3d(x01, y01, 0)



                            Line1_extended.EndPoint = New Point3d(x02_extended, y02_extended, 0)
                           


                            Dim Bearing3_4 As Double = GET_Bearing_rad(x03, y03, x04, y04)
                            Dim Bear_for_1 As Double = Bearing3_4 + PI / 2
                            If Bear_for_1 > 2 * PI Then Bear_for_1 = Bear_for_1 - (2 * PI)


                            Dim circle1 As New Autodesk.AutoCAD.DatabaseServices.Circle
                            circle1.Center = New Point3d(x03 + R1 * Cos(Bear_for_1), y03 + R1 * Sin(Bear_for_1), 0)

                            circle1.Radius = R1
                          

                            Dim Point11 As New Point3dCollection
                            Dim point11_down As New Point3dCollection

                            Dim Deseneaza_in_jos As Boolean = False

                            Line1_extended.IntersectWith(circle1, Intersect.OnBothOperands, Point11, IntPtr.Zero, IntPtr.Zero)

                            If Point11.Count = 1 Then
                                x02 = Point11(0).X
                                y02 = Point11(0).Y
                            ElseIf Point11.Count = 2 Then
                                ' Dim Bear_temp1 As Double = GET_Bearing_rad(x03, y03, Point11(0).X, Point11(0).Y)
                                'Dim Bear_temp2 As Double = GET_Bearing_rad(x03, y03, Point11(1).X, Point11(1).Y)
                                x02 = Point11(0).X
                                y02 = Point11(0).Y


                            ElseIf Point11.Count = 0 Then
                                Deseneaza_in_jos = True
                                Dim Bear_for_1_down As Double = Bear_for_1 + PI
                                If Bear_for_1_down > 2 * PI Then Bear_for_1_down = Bear_for_1_down - 2 * PI

                                Dim circle1_down As New Autodesk.AutoCAD.DatabaseServices.Circle
                                circle1_down.Center = New Point3d(x03 + R1 * Cos(Bear_for_1_down), y03 + R1 * Sin(Bear_for_1_down), 0)
                                circle1_down.Radius = R1
                               

                                Line1_extended.IntersectWith(circle1_down, Intersect.OnBothOperands, point11_down, IntPtr.Zero, IntPtr.Zero)

                                If point11_down.Count = 1 Then
                                    x02 = point11_down(0).X
                                    y02 = point11_down(0).Y
                                ElseIf point11_down.Count = 2 Then
                                    x02 = point11_down(0).X
                                    y02 = point11_down(0).Y
                                ElseIf point11_down.Count = 0 Then

                                    MsgBox("Please verify your radius")
                                    Exit Sub



                                End If ' asta e de la  point11_down.Count = 0


                            End If ' asta e de la Point11.Count = 0





                            Dim Alpha2 As Double = GET_Bearing_rad(x03, y03, x04, y04) + PI / 2
                            If Deseneaza_in_jos = True Then Alpha2 = GET_Bearing_rad(x03, y03, x04, y04) - PI / 2


                            X9 = x03 + R1 * Cos(Alpha2)
                            Y9 = y03 + R1 * Sin(Alpha2)

                            Dim Bear1 As Double = GET_Bearing_rad(x03, y03, x02, y02) + PI / 2
                            If Deseneaza_in_jos = True Then Bear1 = GET_Bearing_rad(x03, y03, x02, y02) - PI / 2

                            If Bear1 < 0 Then Bear1 = Bear1 + 2 * PI


                            Y8 = Y9 + R1 * Sin(Bear1)
                            X8 = X9 + R1 * Cos(Bear1)

                            Dim Xbear, Ybear As Double


                            Ybear = Y8 + 1000 * Sin(GET_Bearing_rad(x03, y03, x02, y02))
                            Xbear = X8 + 1000 * Cos(GET_Bearing_rad(x03, y03, x02, y02))

                            Dim Xint, Yint As Double

                            Xint = Creeaza_Intersectie_X(x01, y01, x02, y02, X8, Y8, Xbear, Ybear)
                            Yint = Creeaza_Intersectie_y(x01, y01, x02, y02, X8, Y8, Xbear, Ybear)

                            Dim Bear_line1 As Double = GET_Bearing_rad(x02, y02, x01, y01)
                            Dim Bear_line2 As Double = GET_Bearing_rad(x02, y02, x03, y03)



                            Dim Unghi12 As Double = Abs(Bear_line1 - Bear_line2)
                            If Unghi12 > PI Then
                                Unghi12 = 2 * PI - Unghi12
                            End If
                            Dim Alpha3 As Double = Unghi12 / 2
                            Dim DistX As Double = R2 / Tan(Alpha3)




                            X5 = Xint + DistX * Cos(Bear_line1)
                            Y5 = Yint + DistX * Sin(Bear_line1)
                            X7 = Xint + DistX * Cos(Bear_line2)
                            Y7 = Yint + DistX * Sin(Bear_line2)

                            X6 = X5 + R2 * Cos(Bear_line1 + PI / 2)
                            Y6 = Y5 + R2 * Sin(Bear_line1 + PI / 2)

                            If Deseneaza_in_jos = True Then
                                X6 = X5 + R2 * Cos(Bear_line1 - PI / 2)
                                Y6 = Y5 + R2 * Sin(Bear_line1 - PI / 2)
                            End If



                            'BULGE CALCS
                            ' BULGE = TAN(Unghi/4)
                            Dim Bear65 As Double = GET_Bearing_rad(X6, Y6, X5, Y5)
                            Dim Bear67 As Double = GET_Bearing_rad(X6, Y6, X7, Y7)
                            Dim Bear903 As Double = GET_Bearing_rad(X9, Y9, x03, y03)
                            Dim Bear98 As Double = GET_Bearing_rad(X9, Y9, X8, Y8)



                            Dim Alpha765 As Double = Abs(Bear65 - Bear67)

                            If Alpha765 > PI Then
                                Alpha765 = 2 * PI - Alpha765
                            End If
                            If Alpha765 > PI Then Alpha765 = PI - Alpha765
                            Unghi_arc2 = Round(Alpha765 * 180 / PI, 4)

                            Dim Alpha0398 As Double = Abs(Bear903 - Bear98)
                            If Alpha0398 > PI Then
                                Alpha0398 = 2 * PI - Alpha0398
                            End If
                            If Alpha0398 > PI Then Alpha0398 = PI - Alpha765
                            Unghi_arc1 = Round(Alpha0398 * 180 / PI, 4)

                            Dim Bulge3 As Double = Tan(Alpha765 / 4)
                            Dim Bulge1 As Double = Tan(Alpha0398 / 4)

                            If Deseneaza_in_jos = True Then
                                Bulge3 = -Bulge3
                                Bulge1 = -Bulge1
                            End If

                            Dim Layer_pipe As String = ComboBox_layer_CL.Text



                            Dim Poly_fillet As New Autodesk.AutoCAD.DatabaseServices.Polyline
                            Poly_fillet.Layer = Layer_pipe


                            Poly_fillet.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(x03, y03), -Bulge1, 0, 0)
                            Poly_fillet.AddVertexAt(1, New Autodesk.AutoCAD.Geometry.Point2d(X8, Y8), 0, 0, 0)
                            Poly_fillet.AddVertexAt(2, New Autodesk.AutoCAD.Geometry.Point2d(X7, Y7), Bulge3, 0, 0)
                            Poly_fillet.AddVertexAt(3, New Autodesk.AutoCAD.Geometry.Point2d(X5, Y5), 0, 0, 0)

                            BTrecord.AppendEntity(Poly_fillet)
                            Trans1.AddNewlyCreatedDBObject(Poly_fillet, True)

                            If CheckBox_draw_od.Checked = True Then
                                Dim Layer_od As String = ComboBox_layer_OD.Text

                                Dim Number_string As String
                                Number_string = Replace(ComboBox_nps.Text, "NPS ", "")
                                If IsNumeric(Number_string) = True And IsNumeric(TextBox_diameter_X.Text) = True Then
                                    Dim NR As Integer = CInt(Number_string)
                                    Dim Dx As Double = CDbl(TextBox_diameter_X.Text)
                                    Dim Off1 As Double = Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000
                                    Dim Obj_coll1 As DBObjectCollection = Poly_fillet.GetOffsetCurves(Off1)
                                    Dim Obj_coll2 As DBObjectCollection = Poly_fillet.GetOffsetCurves(-Off1)
                                    For Each acEnt As Entity In Obj_coll1
                                        acEnt.Layer = Layer_od
                                        '' Add each offset object
                                        BTrecord.AppendEntity(acEnt)

                                        Trans1.AddNewlyCreatedDBObject(acEnt, True)
                                    Next
                                    For Each acEnt As Entity In Obj_coll2
                                        acEnt.Layer = Layer_od
                                        '' Add each offset object
                                        BTrecord.AppendEntity(acEnt)
                                        Trans1.AddNewlyCreatedDBObject(acEnt, True)
                                    Next
                                End If
                            End If









                            Dim Poly_test As New Autodesk.AutoCAD.DatabaseServices.Polyline




                            Poly_test.AddVertexAt(0, New Autodesk.AutoCAD.Geometry.Point2d(x03, y03), 0, 0, 0)
                            Poly_test.AddVertexAt(1, New Autodesk.AutoCAD.Geometry.Point2d(X9, Y9), 0, 0, 0)
                            Poly_test.AddVertexAt(2, New Autodesk.AutoCAD.Geometry.Point2d(X8, Y8), 0, 0, 0)
                            Poly_test.AddVertexAt(3, New Autodesk.AutoCAD.Geometry.Point2d(X7, Y7), 0, 0, 0)
                            Poly_test.AddVertexAt(4, New Autodesk.AutoCAD.Geometry.Point2d(X6, Y6), 0, 0, 0)
                            Poly_test.AddVertexAt(5, New Autodesk.AutoCAD.Geometry.Point2d(X5, Y5), 0, 0, 0)

                            Lungime_linie = Round(GET_distanta_Double_XY(X7, Y7, X8, Y8), 2)
                            Lungime_arc1 = Round(Unghi_arc1 * R1 * PI / 180, 2)
                            Lungime_arc2 = Round(Unghi_arc2 * R2 * PI / 180, 2)

                        End If ' ASTA E DE LA If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line



                        Editor1.Regen()
                        Trans1.Commit()




                    End Using


                End If ' asta e de la Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK

                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
            End Using 'asta e de la lock1
        Catch ex As Exception

            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox_nps_TabIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_nps.TextChanged, TextBox_diameter_X.TextChanged
        Try
            Dim Number_string As String
            Number_string = Replace(ComboBox_nps.Text, "NPS ", "")
            If IsNumeric(Number_string) = True And IsNumeric(TextBox_diameter_X.Text) = True Then
                Dim NR As Integer = CInt(Number_string)
                Dim Dx As Double = CDbl(TextBox_diameter_X.Text)
                TextBox_radius1.Text = Dx * 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000
                TextBox_radius2.Text = Dx * 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000
                TextBox_report.Text = ComboBox_nps.Text & vbCrLf & "Diameter = " & 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000 & " m" _
              & vbCrLf & "Radius = " & Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(NR) / 1000 & " m"
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function Creeaza_Intersectie_X(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double) As Double
        Dim M1, M2 As Double
        If x1 <> 0 And x2 <> 0 And x3 <> 0 And x4 <> 0 And y1 <> 0 And y2 <> 0 And y3 <> 0 And y4 <> 0 Then
            M1 = (y2 - y1) / (x2 - x1)
            M2 = (y4 - y3) / (x4 - x3)

            Return (M1 * x1 - M2 * x3 + y3 - y1) / (M1 - M2)

        Else
            Return 0


        End If

    End Function
    Public Function Creeaza_Intersectie_y(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double) As Double
        Dim M1, M2 As Double
        If x1 <> 0 And x2 <> 0 And x3 <> 0 And x4 <> 0 And y1 <> 0 And y2 <> 0 And y3 <> 0 And y4 <> 0 Then
            M1 = (y2 - y1) / (x2 - x1)
            M2 = (y4 - y3) / (x4 - x3)

            Return y1 + M1 * ((M1 * x1 - M2 * x3 + y3 - y1) / (M1 - M2) - x1)

        Else
            Return 0


        End If

    End Function


    Private Sub Panel2_Click(sender As Object, e As EventArgs) Handles Panel2.Click
        Incarca_existing_layers_to_combobox(ComboBox_layer_CL)
        Incarca_existing_layers_to_combobox(ComboBox_layer_OD)
    End Sub
End Class