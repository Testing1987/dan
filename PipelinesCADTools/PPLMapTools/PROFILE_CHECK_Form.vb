Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class PROFILE_CHECK_Form
    Dim Colectie_butoane As New Specialized.StringCollection
    Dim Freeze_operations As Boolean = False

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button_scale_PICK.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ascunde_butoanele_pentru_forms(Me, Colectie_butoane)
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Using lock1 As DocumentLock = ThisDrawing.LockDocument
            Try
                Dim Labeled_horiz_distance As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Type in the labeled horizontal distance:")
                Dim Horiz_dist_labeled As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Labeled_horiz_distance)

                If Horiz_dist_labeled.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                Dim Point_start_horiz1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                Dim Point_start_horiz_opt1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Pick first point for the horizontal distance:")
                Point_start_horiz1 = Editor1.GetPoint(Point_start_horiz_opt1)
                If Point_start_horiz1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                Dim Point_end_horiz2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                Dim Point_start_horiz_opt2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Pick second point for the horizontal distance:")
                Point_start_horiz_opt2.BasePoint = Point_start_horiz1.Value
                Point_start_horiz_opt2.UseBasePoint = True
                Point_end_horiz2 = Editor1.GetPoint(Point_start_horiz_opt2)
                If Point_end_horiz2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                Dim Labeled_vertical_distance As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Type in the labeled vertical distance:")
                Dim Vertical_dist_labeled As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Labeled_vertical_distance)

                If Vertical_dist_labeled.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                Dim Point_start_vertical1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                Dim Point_start_vertical_opt1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Pick first point for the vertical distance:")
                Point_start_vertical1 = Editor1.GetPoint(Point_start_vertical_opt1)
                If Point_start_vertical1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                Dim Point_end_vertical2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim Point_start_vertical_opt2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Pick second point for the vertical distance:")
                Point_start_vertical_opt2.BasePoint = Point_start_vertical1.Value
                Point_start_vertical_opt2.UseBasePoint = True
                Point_end_vertical2 = Editor1.GetPoint(Point_start_vertical_opt2)
                If Point_end_vertical2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                Dim DELTA_VERTICAL As Double
                DELTA_VERTICAL = Abs(Point_start_vertical1.Value.Y - Point_end_vertical2.Value.Y)
                Dim DELTA_HORIZONTAL As Double
                DELTA_HORIZONTAL = Abs(Point_start_horiz1.Value.X - Point_end_horiz2.Value.X)

                Dim Vertical_exag As Double = DELTA_VERTICAL / Vertical_dist_labeled.Value
                Dim Horizontal_exag As Double = DELTA_HORIZONTAL / Horiz_dist_labeled.Value

                TextBox_horiz_scale.Text = Round(1000 / Horizontal_exag, 4)
                TextBox_vert_scale.Text = Round(1000 / Vertical_exag, 4)



                afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)

                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")

            Catch ex As Exception
                afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                MsgBox(ex.Message)
                Editor1.SetImpliedSelection(Empty_array)
            End Try
        End Using



    End Sub

    Private Sub Button_LABEL_POSITION_Click(sender As System.Object, e As System.EventArgs) Handles Button_LABEL_POSITION.Click
        If Freeze_operations = False Then


            Try
                Freeze_operations = True
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor
                Dim Empty_array() As ObjectId
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    If IsNumeric(TextBox_vert_scale.Text) = True And IsNumeric(TextBox_vert_scale.Text) = True And IsNumeric(TextBox_arrrow_size.Text) = True And IsNumeric(TextBox_X.Text) = True _
                        And IsNumeric(TextBox_Y.Text) = True And IsNumeric(TextBox_gap.Text) = True And IsNumeric(TextBox_dog_length.Text) = True And IsNumeric(TextBox_text_size.Text) = True Then


                        Dim Vertical_exag As Double = 1000 / TextBox_vert_scale.Text
                        Dim Horizontal_exag As Double = 1000 / TextBox_horiz_scale.Text


                        Dim Rezultat_hor As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_hor As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_hor.MessageForAdding = vbLf & "Select a known vertical line (CHAINAGE) and the label for it:" & vbCrLf & "If your selection contains only a line means you selected 0+000 station"

                        Object_prompt_hor.SingleOnly = False
                        Rezultat_hor = Editor1.GetSelection(Object_prompt_hor)


                        Dim Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_vert As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_vert.MessageForAdding = vbLf & "Select a known horizontal line (ELEVATION) and the label for it:"

                        Object_prompt_vert.SingleOnly = False
                        Rezultat_vert = Editor1.GetSelection(Object_prompt_vert)

                        Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem


                        If Rezultat_hor.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_vert.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                            Creaza_layer("NO PLOT", 40, "", False)


123:

                            Dim Descriptia As String = ""
                            Dim Prefix_Elev As String = ""
                            Dim Prefix_chainage As String = ""

                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select required position:")
                                PP1.AllowNone = True
                                Point1 = Editor1.GetPoint(PP1)
                                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Dim Elevatia_cunoscuta As Double = -100000
                                Dim Distanta_de_la_zero1 As Double = -100000
                                Dim Chainage_cunoscuta As Double = -100000

                                Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                                Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                                Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                                Dim polyLinia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline

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

                                If Rezultat_hor.Value.Count > 1 Then
                                    Obj4 = Rezultat_hor.Value.Item(1)
                                    Ent4 = Obj4.ObjectId.GetObject(OpenMode.ForRead)
                                End If


                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    mText_cunoscut = Ent1
                                    If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                                End If

                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    mText_cunoscut = Ent2
                                    If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                                End If

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Text_cunoscut = Ent1
                                    If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                                End If

                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Text_cunoscut = Ent2
                                    If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                                End If

                                If Elevatia_cunoscuta = -100000 Then
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Freeze_operations = False
                                    Exit Sub
                                End If
                                Dim x01, y01, x02, y02, dist1, a1 As Double

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent1
                                    x01 = Linia_cunoscuta.StartPoint.X
                                    y01 = Linia_cunoscuta.StartPoint.Y
                                    x02 = Linia_cunoscuta.EndPoint.X
                                    y02 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.Value.X - x01) ^ 2 + (Point1.Value.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.Value.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If


                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent2
                                    x01 = Linia_cunoscuta.StartPoint.X
                                    y01 = Linia_cunoscuta.StartPoint.Y
                                    x02 = Linia_cunoscuta.EndPoint.X
                                    y02 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.Value.X - x01) ^ 2 + (Point1.Value.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.Value.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent1

                                    x01 = polyLinia_cunoscuta.StartPoint.X
                                    y01 = polyLinia_cunoscuta.StartPoint.Y
                                    x02 = polyLinia_cunoscuta.EndPoint.X
                                    y02 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.Value.X - x01) ^ 2 + (Point1.Value.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.Value.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If

                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent2

                                    x01 = polyLinia_cunoscuta.StartPoint.X
                                    y01 = polyLinia_cunoscuta.StartPoint.Y
                                    x02 = polyLinia_cunoscuta.EndPoint.X
                                    y02 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(y01 - y02) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    dist1 = ((Point1.Value.X - x01) ^ 2 + (Point1.Value.Y - y01) ^ 2) ^ 0.5
                                    a1 = Abs(Point1.Value.X - x01)
                                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                                End If

                                If Distanta_de_la_zero1 = -100000 Then
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Distanta_de_la_zero1 = Distanta_de_la_zero1 / Vertical_exag



                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    mText_cunoscut = Ent3
                                    Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    mText_cunoscut = Ent4
                                    Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Text_cunoscut = Ent3
                                    Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Text_cunoscut = Ent4
                                    Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                                End If

                                If Not Distanta_de_la_zero1 = -100000 Then
                                    If Rezultat_hor.Value.Count = 1 Then Chainage_cunoscuta = 0
                                End If


                                If Chainage_cunoscuta = -100000 Then
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Exit Sub
                                End If
                                Dim x03, y03, x04, y04 As Double
                                Dim Chainage_at_point As Double

                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent3
                                    x03 = Linia_cunoscuta.StartPoint.X
                                    y03 = Linia_cunoscuta.StartPoint.Y
                                    x04 = Linia_cunoscuta.EndPoint.X
                                    y04 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.Value.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    End If
                                End If


                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Linia_cunoscuta = Ent4
                                    x03 = Linia_cunoscuta.StartPoint.X
                                    y03 = Linia_cunoscuta.StartPoint.Y
                                    x04 = Linia_cunoscuta.EndPoint.X
                                    y04 = Linia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.Value.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    End If


                                End If

                                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent3
                                    x03 = polyLinia_cunoscuta.StartPoint.X
                                    y03 = polyLinia_cunoscuta.StartPoint.Y
                                    x04 = polyLinia_cunoscuta.EndPoint.X
                                    y04 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.Value.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    End If
                                End If

                                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    polyLinia_cunoscuta = Ent4
                                    x03 = polyLinia_cunoscuta.StartPoint.X
                                    y03 = polyLinia_cunoscuta.StartPoint.Y
                                    x04 = polyLinia_cunoscuta.EndPoint.X
                                    y04 = polyLinia_cunoscuta.EndPoint.Y
                                    If Abs(x03 - x04) > 0.001 Then
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Freeze_operations = False
                                        Exit Sub
                                    End If
                                    If Point1.Value.X < x03 Then
                                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    Else
                                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.Value.X) / Horizontal_exag
                                    End If
                                End If



                                If Not TextBox_elevation_prefix.Text = "" Then
                                    Prefix_Elev = TextBox_elevation_prefix.Text & " "
                                End If


                                If Not TextBox_chainage_prefix.Text = "" Then
                                    Prefix_chainage = TextBox_chainage_prefix.Text & " "
                                End If

                                If Point1.Value.Y < y01 Then

                                    If Not Prefix_chainage = "" Then
                                        If Not Prefix_Elev = "" Then
                                            Descriptia = Prefix_chainage & Get_chainage_from_double(Chainage_at_point, 2) & vbCrLf & Prefix_Elev & (Get_String_Rounded(Elevatia_cunoscuta - Distanta_de_la_zero1, 2)).ToString
                                        Else
                                            Descriptia = Prefix_chainage & Get_chainage_from_double(Chainage_at_point, 2)
                                        End If
                                    Else
                                        If Not Prefix_Elev = "" Then
                                            Descriptia = Prefix_Elev & (Get_String_Rounded(Elevatia_cunoscuta - Distanta_de_la_zero1, 2)).ToString
                                        End If

                                    End If

                                Else
                                    If Not Prefix_chainage = "" Then
                                        If Not Prefix_Elev = "" Then
                                            Descriptia = Prefix_chainage & Get_chainage_from_double(Chainage_at_point, 2) & vbCrLf & Prefix_Elev & (Get_String_Rounded(Elevatia_cunoscuta + Distanta_de_la_zero1, 2)).ToString
                                        Else
                                            Descriptia = Prefix_chainage & Get_chainage_from_double(Chainage_at_point, 2)
                                        End If
                                    Else
                                        If Not Prefix_Elev = "" Then
                                            Descriptia = Prefix_Elev & (Get_String_Rounded(Elevatia_cunoscuta + Distanta_de_la_zero1, 2)).ToString
                                        End If

                                    End If

                                End If


                                If Not Descriptia = "" Then
                                    Dim Mtext_start As New MText
                                    Mtext_start.Contents = Descriptia
                                    Mtext_start.TextHeight = CDbl(TextBox_text_size.Text)
                                    Mtext_start.ColorIndex = 256
                                    Dim Mleader1 As New MLeader
                                    Dim Nr1 As Integer = Mleader1.AddLeader()
                                    Dim Nr2 As Integer = Mleader1.AddLeaderLine(Nr1)
                                    Mleader1.AddFirstVertex(Nr2, Point1.Value.TransformBy(Curent_UCS))
                                    Mleader1.AddLastVertex(Nr2, New Point3d((Point1.Value).TransformBy(Curent_UCS).X + CDbl(TextBox_X.Text),
                                                                            Point1.Value.TransformBy(Curent_UCS).Y + CDbl(TextBox_Y.Text), 0))

                                    Mleader1.ContentType = ContentType.MTextContent
                                    Mleader1.MText = Mtext_start
                                    Mleader1.Layer = "NO PLOT"
                                    Mleader1.LandingGap = CDbl(TextBox_gap.Text)
                                    Mleader1.ArrowSize = CDbl(TextBox_arrrow_size.Text)
                                    Mleader1.DoglegLength = CDbl(TextBox_dog_length.Text)
                                    Mleader1.Linetype = "BYLAYER"
                                    Mleader1.LineWeight = LineWeight.ByLayer

                                    BTrecord.AppendEntity(Mleader1)
                                    Trans1.AddNewlyCreatedDBObject(Mleader1, True)
                                    Trans1.Commit()
                                End If
                                GoTo 123
                            End Using
                            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Else

                            Editor1.SetImpliedSelection(Empty_array)
                            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End If
                End Using



                Editor1.SetImpliedSelection(Empty_array)

            Catch ex As Exception

                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub



    Private Sub Button_transfer_XL_Click(sender As Object, e As EventArgs) Handles Button_transfer_XL.Click
        If Freeze_operations = False Then


            Try
                Freeze_operations = True
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

                Dim Empty_array() As ObjectId
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    If IsNumeric(TextBox_vert_scale.Text) = True And IsNumeric(TextBox_vert_scale.Text) = True And IsNumeric(TextBox_arrrow_size.Text) = True And IsNumeric(TextBox_X.Text) = True _
                        And IsNumeric(TextBox_Y.Text) = True And IsNumeric(TextBox_gap.Text) = True And IsNumeric(TextBox_dog_length.Text) = True And IsNumeric(TextBox_text_size.Text) = True Then


                        Dim Vertical_exag As Double = 1000 / TextBox_vert_scale.Text
                        Dim Horizontal_exag As Double = 1000 / TextBox_horiz_scale.Text


                        Dim Rezultat_hor As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_hor As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_hor.MessageForAdding = vbLf & "Select a known vertical line (Station) and the label for it:" & vbCrLf & "If your selection contains only a line means you selected 0+000 station"

                        Object_prompt_hor.SingleOnly = False
                        Rezultat_hor = Editor1.GetSelection(Object_prompt_hor)


                        Dim Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_prompt_vert As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_prompt_vert.MessageForAdding = vbLf & "Select a known horizontal line (ELEVATION) and the label for it:"

                        Object_prompt_vert.SingleOnly = False
                        Rezultat_vert = Editor1.GetSelection(Object_prompt_vert)

                        Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem


                        If Rezultat_hor.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_vert.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                            Creaza_layer("NO PLOT", 40, "", False)


                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                                Dim Rezultat_polyHDD As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                                Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the HDD polyline or spline:")

                                Object_Prompt1.SetRejectMessage(vbLf & "Please select a Polyline or spline")
                                Object_Prompt1.AddAllowedClass(GetType(Polyline), True)
                                Object_Prompt1.AddAllowedClass(GetType(Spline), True)

                                Rezultat_polyHDD = Editor1.GetEntity(Object_Prompt1)

                                If Not Rezultat_polyHDD.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    MsgBox("NO HDD")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Exit Sub
                                End If

                                Dim Rezultat_polyGround As Autodesk.AutoCAD.EditorInput.PromptEntityResult

                                Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the Ground polyline:")

                                Object_Prompt2.SetRejectMessage(vbLf & "Please select a Polyline")
                                Object_Prompt2.AddAllowedClass(GetType(Polyline), True)


                                Rezultat_polyGround = Editor1.GetEntity(Object_Prompt2)


                                If Not Rezultat_polyGround.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    MsgBox("NO Ground")
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Exit Sub
                                End If



                                Dim RezultatWater As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                                Object_Prompt.MessageForAdding = vbLf & "Select water lines:"

                                Object_Prompt.SingleOnly = False

                                RezultatWater = Editor1.GetSelection(Object_Prompt)


                                If RezultatWater.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Editor1.WriteMessage(vbLf & "Command:")
                                    MsgBox("NO Water")
                                    Freeze_operations = False
                                    Editor1.SetImpliedSelection(Empty_array)
                                    Exit Sub
                                End If
                                Dim Spacing As Double = 50
                                If IsNumeric(TextBox_spacing.Text) = True Then
                                    Spacing = Abs(CDbl(TextBox_spacing.Text))
                                End If

                                Dim Ent_hdd As Entity = Trans1.GetObject(Rezultat_polyHDD.ObjectId, OpenMode.ForRead)

                                Dim Table1 As New System.Data.DataTable
                                Table1.Columns.Add("STATION", GetType(Double))
                                Table1.Columns.Add("GROUND_ELEVATION", GetType(Double))
                                Table1.Columns.Add("WATER_ELEVATION", GetType(Double))
                                Table1.Columns.Add("HDD_ELEVATION", GetType(Double))

                                If TypeOf Ent_hdd Is Spline Or TypeOf Ent_hdd Is Polyline Then
                                    Dim Spline_hdd As Spline = TryCast(Ent_hdd, Spline)
                                    Dim Poly_hdd As Polyline = TryCast(Ent_hdd, Polyline)
                                    Dim Curve_hdd As Curve = Ent_hdd
                                    Dim Length1 As Double = 0

                                    If IsNothing(Poly_hdd) = True Then
                                        Dim Param_start As Double = Spline_hdd.StartParam
                                        Dim Param_end As Double = Spline_hdd.EndParam
                                        Length1 = Abs(Spline_hdd.GetDistanceAtParameter(Param_start) - Spline_hdd.GetDistanceAtParameter(Param_end))
                                    Else
                                        Length1 = Poly_hdd.Length
                                    End If




                                    Dim Point_no As Integer = 0

                                    If Length1 > Spacing Then
                                        Point_no = CInt(Floor(Length1 / Spacing))
                                    End If

                                    Dim D1 As Double = 0
                                    Dim Poly_ground As Polyline = Trans1.GetObject(Rezultat_polyGround.ObjectId, OpenMode.ForRead)


                                    For i = 0 To Point_no + 1
                                        Dim pt_hdd As Point3d

                                        If D1 <= Length1 Then
                                            pt_hdd = Curve_hdd.GetPointAtDist(D1)
                                        Else
                                            pt_hdd = Curve_hdd.EndPoint
                                        End If


                                        Dim Line1 As New Line(New Point3d(pt_hdd.X, -1000000000, 0), New Point3d(pt_hdd.X, 1000000000, 0))

                                        Dim Col_ground As New Point3dCollection
                                        Col_ground = Intersect_on_both_operands(Line1, Poly_ground)


                                        Dim Col_water_tot As New Point3dCollection

                                        For j = 0 To RezultatWater.Value.Count - 1
                                            Dim Col_water As New Point3dCollection
                                            Dim Curve_water As Curve = Trans1.GetObject(RezultatWater.Value(j).ObjectId, OpenMode.ForRead)
                                            Col_water = Intersect_on_both_operands(Line1, Curve_water)
                                            If Col_water.Count > 0 Then
                                                For k = 0 To Col_water.Count - 1
                                                    Col_water_tot.Add(Col_water(k))
                                                Next
                                            End If
                                        Next







                                        Dim Col_HDD As New Point3dCollection
                                        Col_HDD.Add(pt_hdd)

                                        Populate_table(Trans1, Col_ground, Col_water_tot, Col_HDD, Table1, Rezultat_hor, Rezultat_vert, Horizontal_exag, Vertical_exag)

                                        Line1.Layer = "NO PLOT"
                                        BTrecord.AppendEntity(Line1)
                                        Trans1.AddNewlyCreatedDBObject(Line1, True)

                                        D1 = D1 + Spacing
                                    Next


                                End If

                                Trans1.Commit()

                                If Table1.Rows.Count > 0 Then
                                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                                    'W1.Cells.NumberFormat = "@"

                                    Dim maxRows As Integer = Table1.Rows.Count
                                    Dim maxCols As Integer = Table1.Columns.Count

                                    Dim range As Microsoft.Office.Interop.Excel.Range = W1.Range(W1.Cells(2, 1), W1.Cells(maxRows + 1, maxCols))

                                    Dim values(maxRows, maxCols) As Object

                                    For row = 0 To maxRows - 1
                                        For col = 0 To maxCols - 1
                                            If IsDBNull(Table1.Rows(row).Item(col)) = False Then
                                                values(row, col) = Table1.Rows(row).Item(col)
                                            End If
                                        Next
                                    Next

                                    range.Value2 = values

                                    For i = 0 To Table1.Columns.Count - 1
                                        W1.Cells(1, i + 1).value2 = Table1.Columns(i).ColumnName
                                    Next
                                End If







                            End Using
                            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        Else

                            Editor1.SetImpliedSelection(Empty_array)
                            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End If
                End Using



                Editor1.SetImpliedSelection(Empty_array)

            Catch ex As Exception

                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Public Function Populate_table(ByVal Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction, _
                                ByVal Col_ground As Point3dCollection, _
                                ByVal Col_water As Point3dCollection, _
                                ByVal Col_HDD As Point3dCollection, _
                                ByVal Table1 As System.Data.DataTable, _
                                ByVal Rezultat_hor As Autodesk.AutoCAD.EditorInput.PromptSelectionResult, _
                                ByVal Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult, _
                                ByVal Horizontal_exag As Double, _
                                ByVal Vertical_exag As Double)

        Dim Empty_array() As ObjectId
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        If Col_ground.Count > 0 Then
            For j = 0 To Col_ground.Count - 1
                Table1.Rows.Add()
                Dim Point1 As New Point3d
                Point1 = Col_ground(j)

                Dim PointW1 As New Point3d
                If Col_water.Count > 0 Then
                    PointW1 = Col_water(0)
                End If


                Dim PointH1 As New Point3d
                PointH1 = Col_HDD(0)

                Dim Elevatia_cunoscuta As Double = -100000
                Dim Elevation_dist1 As Double = -100000
                Dim Elevation_distW1 As Double = -100000
                Dim Elevation_distH1 As Double = -100000
                Dim Chainage_cunoscuta As Double = -100000

                Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                Dim polyLinia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline

                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj1 = Rezultat_vert.Value.Item(0)
                Dim Ent1 As Entity
                Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)

                Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj2 = Rezultat_vert.Value.Item(1)
                Dim Ent2 As Entity
                Ent2 = Trans1.GetObject(Obj2.ObjectId, OpenMode.ForRead)

                Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj3 = Rezultat_hor.Value.Item(0)
                Dim Ent3 As Entity
                Ent3 = Trans1.GetObject(Obj3.ObjectId, OpenMode.ForRead)

                Dim Obj4 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Dim Ent4 As Entity

                If Rezultat_hor.Value.Count > 1 Then
                    Obj4 = Rezultat_hor.Value.Item(1)
                    Ent4 = Trans1.GetObject(Obj4.ObjectId, OpenMode.ForRead)
                End If


                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent1
                    If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                End If

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent2
                    If IsNumeric(Replace(mText_cunoscut.Text, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(mText_cunoscut.Text, "'", ""))
                End If

                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent1
                    If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                End If

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent2
                    If IsNumeric(Replace(Text_cunoscut.TextString, "'", "")) = True Then Elevatia_cunoscuta = CDbl(Replace(Text_cunoscut.TextString, "'", ""))
                End If

                If Elevatia_cunoscuta = -100000 Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Freeze_operations = False
                    Exit Function
                End If




                Dim x01, y01, x02, y02, dist1, a1 As Double
                Dim distW1, aW1 As Double
                Dim distH1, aH1 As Double

                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent1
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If

                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point1.X - x01)
                    Elevation_dist1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                    If Col_water.Count > 0 Then
                        distW1 = ((PointW1.X - x01) ^ 2 + (PointW1.Y - y01) ^ 2) ^ 0.5
                        aW1 = Abs(PointW1.X - x01)
                        Elevation_distW1 = (distW1 ^ 2 - aW1 ^ 2) ^ 0.5
                    End If


                    distH1 = ((PointH1.X - x01) ^ 2 + (PointH1.Y - y01) ^ 2) ^ 0.5
                    aH1 = Abs(PointH1.X - x01)
                    Elevation_distH1 = (distH1 ^ 2 - aH1 ^ 2) ^ 0.5

                End If


                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent2
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If
                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point1.X - x01)
                    Elevation_dist1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                    If Col_water.Count > 0 Then
                        distW1 = ((PointW1.X - x01) ^ 2 + (PointW1.Y - y01) ^ 2) ^ 0.5
                        aW1 = Abs(PointW1.X - x01)
                        Elevation_distW1 = (distW1 ^ 2 - aW1 ^ 2) ^ 0.5
                    End If

                    distH1 = ((PointH1.X - x01) ^ 2 + (PointH1.Y - y01) ^ 2) ^ 0.5
                    aH1 = Abs(PointH1.X - x01)
                    Elevation_distH1 = (distH1 ^ 2 - aH1 ^ 2) ^ 0.5

                End If

                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    polyLinia_cunoscuta = Ent1

                    x01 = polyLinia_cunoscuta.StartPoint.X
                    y01 = polyLinia_cunoscuta.StartPoint.Y
                    x02 = polyLinia_cunoscuta.EndPoint.X
                    y02 = polyLinia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If
                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point1.X - x01)
                    Elevation_dist1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                    If Col_water.Count > 0 Then
                        distW1 = ((PointW1.X - x01) ^ 2 + (PointW1.Y - y01) ^ 2) ^ 0.5
                        aW1 = Abs(PointW1.X - x01)
                        Elevation_distW1 = (distW1 ^ 2 - aW1 ^ 2) ^ 0.5
                    End If


                    distH1 = ((PointH1.X - x01) ^ 2 + (PointH1.Y - y01) ^ 2) ^ 0.5
                    aH1 = Abs(PointH1.X - x01)
                    Elevation_distH1 = (distH1 ^ 2 - aH1 ^ 2) ^ 0.5


                End If

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    polyLinia_cunoscuta = Ent2

                    x01 = polyLinia_cunoscuta.StartPoint.X
                    y01 = polyLinia_cunoscuta.StartPoint.Y
                    x02 = polyLinia_cunoscuta.EndPoint.X
                    y02 = polyLinia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If

                    dist1 = ((Point1.X - x01) ^ 2 + (Point1.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point1.X - x01)
                    Elevation_dist1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                    If Col_water.Count > 0 Then
                        distW1 = ((PointW1.X - x01) ^ 2 + (PointW1.Y - y01) ^ 2) ^ 0.5
                        aW1 = Abs(PointW1.X - x01)
                        Elevation_distW1 = (distW1 ^ 2 - aW1 ^ 2) ^ 0.5
                    End If


                    distH1 = ((PointH1.X - x01) ^ 2 + (PointH1.Y - y01) ^ 2) ^ 0.5
                    aH1 = Abs(PointH1.X - x01)
                    Elevation_distH1 = (distH1 ^ 2 - aH1 ^ 2) ^ 0.5


                End If

                If Elevation_dist1 = -100000 Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Freeze_operations = False
                    Exit Function
                End If

                Elevation_dist1 = Elevation_dist1 / Vertical_exag



                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent3
                    Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                End If

                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent4
                    Dim numar_fara_plus As String = Replace(mText_cunoscut.Text, "+", "")
                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent3
                    Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                End If

                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent4
                    Dim numar_fara_plus As String = Replace(Text_cunoscut.TextString, "+", "")
                    If IsNumeric(numar_fara_plus) = True Then Chainage_cunoscuta = CDbl(numar_fara_plus)
                End If

                If Not Elevation_dist1 = -100000 Then
                    If Rezultat_hor.Value.Count = 1 Then Chainage_cunoscuta = 0
                End If


                If Chainage_cunoscuta = -100000 Then
                    Freeze_operations = False
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Function
                End If
                Dim x03, y03, x04, y04 As Double
                Dim Chainage_at_point As Double

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent3
                    x03 = Linia_cunoscuta.StartPoint.X
                    y03 = Linia_cunoscuta.StartPoint.Y
                    x04 = Linia_cunoscuta.EndPoint.X
                    y04 = Linia_cunoscuta.EndPoint.Y
                    If Abs(x03 - x04) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If
                    If Point1.X < x03 Then
                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                    Else
                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                    End If
                End If


                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent4
                    x03 = Linia_cunoscuta.StartPoint.X
                    y03 = Linia_cunoscuta.StartPoint.Y
                    x04 = Linia_cunoscuta.EndPoint.X
                    y04 = Linia_cunoscuta.EndPoint.Y
                    If Abs(x03 - x04) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If
                    If Point1.X < x03 Then
                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                    Else
                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                    End If


                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    polyLinia_cunoscuta = Ent3
                    x03 = polyLinia_cunoscuta.StartPoint.X
                    y03 = polyLinia_cunoscuta.StartPoint.Y
                    x04 = polyLinia_cunoscuta.EndPoint.X
                    y04 = polyLinia_cunoscuta.EndPoint.Y
                    If Abs(x03 - x04) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If
                    If Point1.X < x03 Then
                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                    Else
                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                    End If
                End If

                If TypeOf Ent4 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    polyLinia_cunoscuta = Ent4
                    x03 = polyLinia_cunoscuta.StartPoint.X
                    y03 = polyLinia_cunoscuta.StartPoint.Y
                    x04 = polyLinia_cunoscuta.EndPoint.X
                    y04 = polyLinia_cunoscuta.EndPoint.Y
                    If Abs(x03 - x04) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Freeze_operations = False
                        Exit Function
                    End If
                    If Point1.X < x03 Then
                        Chainage_at_point = Chainage_cunoscuta - Abs(x03 - Point1.X) / Horizontal_exag
                    Else
                        Chainage_at_point = Chainage_cunoscuta + Abs(x03 - Point1.X) / Horizontal_exag
                    End If
                End If




                Table1.Rows(Table1.Rows.Count - 1).Item("STATION") = Chainage_at_point
                If Point1.Y < y01 Then
                    Table1.Rows(Table1.Rows.Count - 1).Item("GROUND_ELEVATION") = Elevatia_cunoscuta - Elevation_dist1
                Else
                    Table1.Rows(Table1.Rows.Count - 1).Item("GROUND_ELEVATION") = Elevatia_cunoscuta + Elevation_dist1
                End If

                If Col_water.Count > 0 Then
                    If PointW1.Y < y01 Then
                        Table1.Rows(Table1.Rows.Count - 1).Item("WATER_ELEVATION") = Elevatia_cunoscuta - Elevation_distW1
                    Else
                        Table1.Rows(Table1.Rows.Count - 1).Item("WATER_ELEVATION") = Elevatia_cunoscuta + Elevation_distW1
                    End If
                End If

                If PointH1.Y < y01 Then
                    Table1.Rows(Table1.Rows.Count - 1).Item("HDDD_ELEVATION") = Elevatia_cunoscuta - Elevation_distH1
                Else
                    Table1.Rows(Table1.Rows.Count - 1).Item("HDD_ELEVATION") = Elevatia_cunoscuta + Elevation_distH1
                End If

            Next
        End If



    End Function


End Class