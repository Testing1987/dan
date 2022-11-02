Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Split_deflection_for_stress_form
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Split_deflection_for_stress_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        ComboBox_NPS.Items.Add(0.125)
        ComboBox_NPS.Items.Add(0.25)
        ComboBox_NPS.Items.Add(0.375)
        ComboBox_NPS.Items.Add(0.5)
        ComboBox_NPS.Items.Add(0.75)
        ComboBox_NPS.Items.Add(10)
        ComboBox_NPS.Items.Add(1.25)
        ComboBox_NPS.Items.Add(1.5)
        ComboBox_NPS.Items.Add(2)
        ComboBox_NPS.Items.Add(2.5)
        ComboBox_NPS.Items.Add(3)
        ComboBox_NPS.Items.Add(3.5)
        ComboBox_NPS.Items.Add(4)
        ComboBox_NPS.Items.Add(5)
        ComboBox_NPS.Items.Add(6)
        ComboBox_NPS.Items.Add(8)
        ComboBox_NPS.Items.Add(10)
        ComboBox_NPS.Items.Add(12)
        ComboBox_NPS.Items.Add(14)
        ComboBox_NPS.Items.Add(16)
        ComboBox_NPS.Items.Add(18)
        ComboBox_NPS.Items.Add(20)
        ComboBox_NPS.Items.Add(22)
        ComboBox_NPS.Items.Add(24)
        ComboBox_NPS.Items.Add(26)
        ComboBox_NPS.Items.Add(28)
        ComboBox_NPS.Items.Add(30)
        ComboBox_NPS.Items.Add(32)
        ComboBox_NPS.Items.Add(34)
        ComboBox_NPS.Items.Add(36)
        ComboBox_NPS.Items.Add(38)
        ComboBox_NPS.Items.Add(40)
        ComboBox_NPS.Items.Add(42)
        ComboBox_NPS.Items.Add(44)
        ComboBox_NPS.Items.Add(46)
        ComboBox_NPS.Items.Add(48)
        ComboBox_NPS.Items.Add(50)
        ComboBox_NPS.Items.Add(52)
        ComboBox_NPS.Items.Add(54)
        ComboBox_NPS.Items.Add(56)
        ComboBox_NPS.Items.Add(58)
        ComboBox_NPS.Items.Add(60)
        ComboBox_NPS.Items.Add(62)
        ComboBox_NPS.Items.Add(64)
        ComboBox_NPS.Items.Add(66)
        ComboBox_NPS.Items.Add(68)
        ComboBox_NPS.Items.Add(70)
        ComboBox_NPS.Items.Add(72)
        ComboBox_NPS.Items.Add(74)
        ComboBox_NPS.Items.Add(76)
        ComboBox_NPS.Items.Add(78)
        ComboBox_NPS.Items.Add(80)

        ComboBox_NPS.SelectedIndex = 26

    End Sub
    Private Sub Button_split_Click(sender As Object, e As EventArgs) Handles Button_split.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        If IsNumeric(TextBox_number_of_splits.Text) = False Then
            MsgBox("Please provide a numeric value for the number of splits")
            Exit Sub
        End If
        If IsNumeric(TextBox_distance.Text) = False Then
            MsgBox("Please provide a numeric value for the joint length")
            Exit Sub
        End If


        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select polyline node:")
                Rezultat1 = Editor1.GetEntity(Object_Prompt)
                If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If

                Dim Poly1 As Polyline
                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim Ent1 As Entity
                            Ent1 = Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForRead)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Poly1 = Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForWrite)
                            Else
                                Editor1.WriteMessage(vbLf & "No Polyline")
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If
                        End Using
                    End If
                End If

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    If IsNothing(Poly1) = False Then
                        Dim Point_on_poly As New Point3d
                        Point_on_poly = Poly1.GetClosestPointTo(Rezultat1.PickedPoint, True)
                        Dim ParamPt0 As Double = Round(Poly1.GetParameterAtPoint(Point_on_poly), 0)

                        Dim Point0 As New Point3d
                        Point0 = Poly1.GetPointAtParameter(ParamPt0)
                        Dim JointL As Double = CDbl(TextBox_distance.Text)

                        Dim vect1 As New Vector3d
                        vect1 = Point0.GetVectorTo(Poly1.GetPointAtParameter(ParamPt0 + 1))
                        Dim vect2 As New Vector3d
                        vect2 = Point0.GetVectorTo(Poly1.GetPointAtParameter(ParamPt0 - 1))
                        Dim Angle As Double = vect1.GetAngleTo(vect2)
                        Dim Nr_split_not_rounded As Double = CDbl(TextBox_number_of_splits.Text)
                        Dim Nr_split As Integer = Round(Nr_split_not_rounded, 0)

                        Dim Distanta0 As Double = Poly1.GetDistanceAtParameter(ParamPt0)

                        If Nr_split > 2 Then
                            Dim circle1 As New Circle(Point0, Vector3d.ZAxis, 2 * JointL)
                            Dim Colint1 As New Point3dCollection
                            circle1.IntersectWith(Poly1, Intersect.OnBothOperands, Colint1, IntPtr.Zero, IntPtr.Zero)
                            If Colint1.Count = 2 Then
                                Dim Pt1 As New Point3d
                                Pt1 = Colint1(0)
                                Dim Pt2 As New Point3d
                                Pt2 = Colint1(1)
                                Dim Linie1 As New Line(Point0, New Point3d((Pt1.X + Pt2.X) / 2, (Pt1.Y + Pt2.Y) / 2, (Pt1.Z + Pt2.Z) / 2))

1025:
                                Dim Linie2 As New Line(Point0, Pt1)
                                Dim Pt_at_jL As Point3d
                                Pt_at_jL = Linie2.GetPointAtDist(JointL)
                                Dim Linie2JL As New Line(Point0, Pt_at_jL)
                                Linie2JL.TransformBy(Matrix3d.Rotation(-(PI - Angle) / Nr_split, Vector3d.ZAxis, Pt_at_jL))
                                Linie2.TransformBy(Matrix3d.Rotation(-(PI - Angle) / Nr_split, Vector3d.ZAxis, Pt1))
                                Dim Colint2 As New Point3dCollection
                                Linie2.IntersectWith(Linie1, Intersect.OnBothOperands, Colint2, IntPtr.Zero, IntPtr.Zero)

                                If Colint2.Count = 0 Then
                                    Pt1 = Colint1(1)
                                    Pt2 = Colint1(0)
                                    GoTo 1025
                                End If

                                Dim Punct1 As New Point3d
                                Punct1 = Linie2JL.EndPoint
                                Dim Punct2 As New Point3d
                                Punct2 = Linie2JL.StartPoint

                                Dim Poly2 As New Polyline
                                Poly2.AddVertexAt(0, New Point2d(Punct1.X, Punct1.Y), 0, 0, 0)
                                Poly2.AddVertexAt(1, New Point2d(Punct2.X, Punct2.Y), 0, 0, 0)

                                For i = 2 To Nr_split - 1
                                    Dim Linie3 As New Line
                                    Linie3 = Linie2JL.Clone
                                    Linie3.TransformBy(Matrix3d.Displacement(Punct1.GetVectorTo(Punct2)))
                                    Linie3.TransformBy(Matrix3d.Rotation(-(PI - Angle) / Nr_split, Vector3d.ZAxis, Punct2))
                                    Poly2.AddVertexAt(i, New Point2d(Linie3.StartPoint.X, Linie3.StartPoint.Y), 0, 0, 0)
                                    Punct1 = Linie3.EndPoint
                                    Punct2 = Linie3.StartPoint
                                    Linie2JL = Linie3.Clone
                                Next

                                Dim Punctm As New Point3d

                                If Nr_split / 2 = Round(Nr_split / 2, 0) Then
                                    Punctm = New Point3d((Poly2.GetPoint3dAt(Nr_split / 2).X + Poly2.GetPoint3dAt(Nr_split / 2 - 1).X) / 2, _
                                                         (Poly2.GetPoint3dAt(Nr_split / 2).Y + Poly2.GetPoint3dAt(Nr_split / 2 - 1).Y) / 2, _
                                                         (Poly2.GetPoint3dAt(Nr_split / 2).Z + Poly2.GetPoint3dAt(Nr_split / 2 - 1).Z) / 2)
                                Else
                                    Punctm = Poly2.GetPoint3dAt((Nr_split - 1) / 2)
                                End If

                                Poly2.TransformBy(Matrix3d.Displacement(Punctm.GetVectorTo(Point0)))
                                Linie1.TransformBy(Matrix3d.Displacement(Point0.GetVectorTo(Poly2.StartPoint)))
                                Dim Colint3 As New Point3dCollection
                                Linie1.IntersectWith(Poly1, Intersect.ExtendThis, Colint3, IntPtr.Zero, IntPtr.Zero)

                                If Colint3.Count > 0 Then
                                    Poly2.TransformBy(Matrix3d.Displacement(Poly2.StartPoint.GetVectorTo(Colint3(0))))
                                    Dim Param_st As Double = Poly1.GetParameterAtPoint(Poly2.StartPoint)
                                    Dim Param_end As Double = Poly1.GetParameterAtPoint(Poly2.EndPoint)
                                    If Param_st > Param_end Then
                                        Poly2.ReverseCurve()
                                    End If
                                    Dim poly11 As Polyline
                                    poly11 = Trans1.GetObject(Poly1.ObjectId, OpenMode.ForWrite)
                                    poly11.RemoveVertexAt(ParamPt0)
                                    For i = Poly2.NumberOfVertices - 1 To 0 Step -1
                                        poly11.AddVertexAt(ParamPt0, Poly2.GetPoint2dAt(i), 0, 0, 0)

                                    Next
                                Else
                                    MsgBox("Joint Length is too long or polyline is too short")
                                End If
                            Else
                                MsgBox("Joint Length is too long or polyline is too short")
                            End If
                        ElseIf Nr_split = 2 Then
                            If Poly1.Length >= Distanta0 + 0.5 * JointL / Sin(Angle / 2) Or Distanta0 - 0.5 * JointL / Sin(Angle / 2) < 0 Then
                                Dim Pt1 As New Point3d
                                Pt1 = Poly1.GetPointAtDist(Distanta0 + 0.5 * JointL / Sin(Angle / 2))
                                Dim Pt2 As New Point3d
                                Pt2 = Poly1.GetPointAtDist(Distanta0 - 0.5 * JointL / Sin(Angle / 2))
                                Dim Linie1 As New Line(Pt1, Pt2)

                                Dim Poly2 As New Polyline
                                Poly2.AddVertexAt(0, New Point2d(Linie1.StartPoint.X, Linie1.StartPoint.Y), 0, 0, 0)
                                Poly2.AddVertexAt(1, New Point2d(Linie1.EndPoint.X, Linie1.EndPoint.Y), 0, 0, 0)

                                Dim Param_st As Double = Poly1.GetParameterAtPoint(Poly2.StartPoint)
                                Dim Param_end As Double = Poly1.GetParameterAtPoint(Poly2.EndPoint)
                                If Param_st > Param_end Then
                                    Poly2.ReverseCurve()
                                End If

                                Dim poly11 As Polyline
                                poly11 = Trans1.GetObject(Poly1.ObjectId, OpenMode.ForWrite)
                                poly11.RemoveVertexAt(ParamPt0)
                                For i = Poly2.NumberOfVertices - 1 To 0 Step -1
                                    poly11.AddVertexAt(ParamPt0, Poly2.GetPoint2dAt(i), 0, 0, 0)
                                Next

                            Else
                                MsgBox("Joint Length is too long or polyline is too short")
                            End If
                        Else
                            MsgBox("Review your angle or joint length")
                        End If
                    End If
                    Trans1.Commit()
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_calculate_MAX_BEND_ANGLE_Click(sender As Object, e As EventArgs) Handles Button_calculate_MAX_BEND_ANGLE.Click
        If IsNumeric(TextBox_tan_length.Text) = False Then
            MsgBox("Please provide a numeric value for the tangent length")
            Exit Sub
        End If
        If IsNumeric(TextBox_joint_length_calcs.Text) = False Then
            MsgBox("Please provide a numeric value for the joint length")
            Exit Sub
        End If
        If IsNumeric(TextBox_degg_per_bend.Text) = False Then
            MsgBox("Please provide a numeric value for the Degrees per bend")
            Exit Sub
        End If

        Dim Diam_m As Double
        If RadioButton_m.Checked = True Then
            Diam_m = 2 * Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(CDbl(ComboBox_NPS.Text)) / 1000
        ElseIf RadioButton_m.Checked = True Then
            Diam_m = CDbl(ComboBox_NPS.Text)
        Else
            Diam_m = CDbl(ComboBox_NPS.Text) / 12
        End If


        Dim JL As Double = CDbl(TextBox_joint_length_calcs.Text)
        Dim TanL As Double = 2 * CDbl(TextBox_tan_length.Text)
        If CheckBox_double_joint.Checked = True Then
            TanL = 2 * TanL
        End If
        TextBox_number_of_bends.Text = 1 / Diam_m * ((JL - TanL))
        TextBox_max_bend_angle.Text = CDbl(TextBox_degg_per_bend.Text) / Diam_m * ((JL - TanL))
    End Sub

    
    Private Sub RadioButton_m_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_m.CheckedChanged
        If RadioButton_m.Checked = True Then
            Label_IM1.Text = "[m]"
            Label_IM2.Text = "[m]"
        ElseIf RadioButton_f.Checked = True Then
            Label_IM1.Text = "[feet]"
            Label_IM2.Text = "[feet]"
        Else
            Label_IM1.Text = "[in]"
            Label_IM2.Text = "[in]"
        End If
    End Sub
    Private Sub RadioButton_i_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_i.CheckedChanged
        If RadioButton_i.Checked = True Then
            Label_IM1.Text = "[in]"
            Label_IM2.Text = "[in]"
        ElseIf RadioButton_f.Checked = True Then
            Label_IM1.Text = "[feet]"
            Label_IM2.Text = "[feet]"
        Else
            Label_IM1.Text = "[m]"
            Label_IM2.Text = "[m]"
        End If
    End Sub
    Private Sub RadioButton_f_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_f.CheckedChanged
        If RadioButton_f.Checked = True Then
            Label_IM1.Text = "[feet]"
            Label_IM2.Text = "[feet]"
        ElseIf RadioButton_i.Checked = True Then
            Label_IM1.Text = "[in]"
            Label_IM2.Text = "[in]"
        Else
            Label_IM1.Text = "[m]"
            Label_IM2.Text = "[m]"
        End If
    End Sub
End Class