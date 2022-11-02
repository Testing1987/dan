Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Survey_band_form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Data_table_survey As System.Data.DataTable

    Private Sub Survey_band_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_table_survey = New System.Data.DataTable
        Data_table_survey.Columns.Add("CHAINAGE", GetType(Double))
        Data_table_survey.Columns.Add("CROSSING_TYPE", GetType(String))
        Data_table_survey.Columns.Add("DESCRIPTION", GetType(String))
        Data_table_survey.Columns.Add("PAGE", GetType(Integer))
        Data_table_survey.Columns.Add("BERMDIST", GetType(Double))
    End Sub

    Private Sub Button_TRANSFER_TO_AUTOCAD_Click(sender As Object, e As EventArgs) Handles Button_TRANSFER_TO_AUTOCAD.Click

        If IsNothing(Data_table_survey) = True Then
            MsgBox("You didn't load any data from excel")
            Exit Sub
        End If
        If Data_table_survey.Rows.Count < 1 Then
            MsgBox("You didn't load any data from excel")
            Exit Sub
        End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Point_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify insertion point")
                PP0.AllowNone = True
                Point_rezult = Editor1.GetPoint(PP0)
                If Not Point_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If


                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        Dim X As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).X
                        Dim y As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).Y
                        Dim z As Double = Point_rezult.Value.TransformBy(curent_ucs_matrix).Z
                        Dim X0 As Double = X

                    Dim PageNR As Integer = 0

                    Dim is_BERM As Boolean = False


                    Dim Start_Hor_line As Point3d
                    Dim BERM_Start As Double
                    Dim Count_berm As Integer

                    Dim BERM_DIST As Double = 0
                    Dim Distanta_intre_coloane As Double = 10
                    Dim Delta As Double = 0

                    For i = 0 To Data_table_survey.Rows.Count - 1

                        Dim Chainage_text As String
                        If IsDBNull(Data_table_survey.Rows(i).Item("CHAINAGE")) = False Then
                            Chainage_text = Get_chainage_from_double(Data_table_survey.Rows(i).Item("CHAINAGE"), 1)
                        Else
                            MsgBox("Chainage problem")
                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                            Exit Sub
                        End If

                        Dim Valoare1 As Double = Round(Data_table_survey.Rows(i).Item("CHAINAGE"), 1)
                        
                        



                        If Not Data_table_survey.Rows(i).Item("CROSSING_TYPE") = "MATCHLINE" Then
                            If IsDBNull(Data_table_survey.Rows(i).Item("PAGE")) = False Then
                                PageNR = Data_table_survey.Rows(i).Item("PAGE")
                            Else
                                MsgBox("PAGE number problem")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If

                            If PageNR > 0 Then
                                Dim CROSSING_TYPE As String
                                If IsDBNull(Data_table_survey.Rows(i).Item("CROSSING_TYPE")) = False Then
                                    CROSSING_TYPE = Data_table_survey.Rows(i).Item("CROSSING_TYPE")
                                Else
                                    MsgBox("Crossing type problem")
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If

                                Dim Description As String


                                If IsDBNull(Data_table_survey.Rows(i).Item("DESCRIPTION")) = False Then
                                    Description = Data_table_survey.Rows(i).Item("DESCRIPTION")
                                    Dim rand() As String = Description.Split(vbCrLf)
                                    If rand.Length > 3 Then
                                        Delta = 5
                                    Else
                                        Delta = 0
                                    End If
                                Else
                                    MsgBox("DESCRIPTION problem")
                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                    Exit Sub
                                End If

                                Select Case CROSSING_TYPE

                                    Case "DEFLECTION"
                                        Dim Poly_l As New Polyline
                                        Poly_l.AddVertexAt(0, New Point2d(X + 2.435 - 2.435, y - 34.38 + 26.472 + 4.16), 0, 0, 0)
                                        Poly_l.AddVertexAt(1, New Point2d(X + 2.435 - 2.435, y - 34.38 + 26.472), 0, 0, 0)
                                        Poly_l.AddVertexAt(2, New Point2d(X + 2.435 - 2.8 - 2.435, y - 34.38 + 26.472 + 4.16), 0, 0, 0)
                                        BTrecord.AppendEntity(Poly_l)
                                        Trans1.AddNewlyCreatedDBObject(Poly_l, True)

                                        Dim Poly_arc As New Polyline
                                        Poly_arc.AddVertexAt(0, New Point2d(X + 2.435 - 3.377 - 2.435, y - 34.38 + 26.472 + 2.446), -Tan((117 * PI / 180) / 4), 0, 0)
                                        Poly_arc.AddVertexAt(1, New Point2d(X + 2.435 + 0.997 - 2.435, y - 34.38 + 26.472 + 3.092), 0, 0, 0)

                                        BTrecord.AppendEntity(Poly_arc)
                                        Trans1.AddNewlyCreatedDBObject(Poly_arc, True)

                                        Dim Description_text As String = Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                        If is_BERM = True Then Count_berm = Count_berm + 1

                                    Case "POWER"

                                        If RadioButton_right_left.Checked = True Then
                                            X = X - 5
                                        Else
                                            X = X + 5
                                        End If

                                        Dim Description_text As String = Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y - 7.908, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                        If RadioButton_right_left.Checked = True Then
                                            X = X - 5
                                        Else
                                            X = X + 5
                                        End If

                                        If is_BERM = True Then Count_berm = Count_berm + 1

                                    Case "SEISMIC"

                                        Dim Description_text As String = "℄ " & Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y - 7.908, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                        If is_BERM = True Then Count_berm = Count_berm + 1


                                    Case "FENCE"

                                        Dim Description_text As String = "℄ " & Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y - 7.908, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                        If is_BERM = True Then Count_berm = Count_berm + 1


                                    Case "CL"

                                        Dim Description_text As String = "℄ " & Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y - 7.908, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                        If is_BERM = True Then Count_berm = Count_berm + 1

                                    Case "MUSKEG"
                                        Dim Description_text As String = Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y - 7.908, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                        If is_BERM = True Then Count_berm = Count_berm + 1

                                    Case "GENERAL"
                                        Dim Rand() As String = Description.Split(vbCrLf)

                                        If Rand.Length > 3 Then
                                            If RadioButton_left_right.Checked = True Then
                                                X = X + 5
                                            Else
                                                X = X - 5
                                            End If
                                        End If
                                       
                                        Dim Description_text As String = Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y - 7.908, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                        If is_BERM = True Then Count_berm = Count_berm + 1

                                    Case "BERM"
                                        Dim Description_text As String = "℄ " & Description
                                        Dim Mtext1 As New MText
                                        Mtext1.Contents = Description_text
                                        Mtext1.TextHeight = 2.5
                                        Mtext1.Rotation = PI / 2
                                        Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                        Mtext1.Location = New Point3d(X, y - 7.908, z)
                                        BTrecord.AppendEntity(Mtext1)
                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)
                                        If is_BERM = True Then
                                            MsgBox("Berm inside berm?!!!!" & vbCrLf & "At " & Chainage_text)
                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                            Exit Sub
                                        End If


                                    Case "BERMS START"
                                        Dim Linie1 As New Line(New Point3d(X, y - 7.908, z), New Point3d(X, y + 30, z))
                                        BTrecord.AppendEntity(Linie1)
                                        Trans1.AddNewlyCreatedDBObject(Linie1, True)


                                        Dim Mtext2 As New MText
                                        Mtext2.Contents = Chainage_text
                                        Mtext2.TextHeight = 2.5
                                        Mtext2.Rotation = PI / 2
                                        Mtext2.Location = New Point3d(X, y - 33, z)
                                        Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                        BTrecord.AppendEntity(Mtext2)
                                        Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                        If IsDBNull(Data_table_survey.Rows(i).Item("BERMDIST")) = False Then
                                            BERM_DIST = Data_table_survey.Rows(i).Item("BERMDIST")
                                        End If

                                        Start_Hor_line = New Point3d(X, y + 27, z)
                                        BERM_Start = CDbl(Replace(Chainage_text, "+", ""))
                                        is_BERM = True

                                    Case "BERMS END"
                                        If is_BERM = True Then

                                            If Count_berm = 1 Then
                                                If RadioButton_right_left.Checked = True Then
                                                    X = X - 10
                                                Else
                                                    X = X + 10
                                                End If
                                            End If
                                            If Count_berm = 0 Then
                                                If RadioButton_right_left.Checked = True Then
                                                    X = X - 20
                                                Else
                                                    X = X + 20
                                                End If
                                            End If

                                            Dim Linie1 As New Line(New Point3d(X, y - 7.908, z), New Point3d(X, y + 30, z))
                                            BTrecord.AppendEntity(Linie1)
                                            Trans1.AddNewlyCreatedDBObject(Linie1, True)

                                            Dim Linie2 As New Line(Start_Hor_line, New Point3d(X, y + 27, z))

                                            Dim Poly1 As New Polyline
                                            Poly1.AddVertexAt(0, New Point2d(Linie2.StartPoint.X, Linie2.StartPoint.Y), 0, 0, 2)


                                            Dim Minus As Integer = 1
                                            If RadioButton_right_left.Checked = True Then
                                                Minus = -1
                                            End If

                                            Poly1.AddVertexAt(1, New Point2d(Linie2.StartPoint.X + 5 * Minus, Linie2.StartPoint.Y), 0, 0, 0)
                                            Poly1.AddVertexAt(2, New Point2d(Linie2.EndPoint.X - 5 * Minus, Linie2.StartPoint.Y), 0, 2, 0)
                                            Poly1.AddVertexAt(3, New Point2d(Linie2.EndPoint.X, Linie2.StartPoint.Y), 0, 0, 0)
                                            Poly1.Elevation = z


                                            BTrecord.AppendEntity(Poly1)
                                            Trans1.AddNewlyCreatedDBObject(Poly1, True)








                                            If IsDBNull(Data_table_survey.Rows(i).Item("BERMDIST")) = False Then
                                                BERM_DIST = Data_table_survey.Rows(i).Item("BERMDIST")
                                            End If


                                            Dim Mtext1_UP As New MText
                                            Mtext1_UP.Contents = Round(((CDbl(Replace(Chainage_text, "+", "")) - BERM_Start) / BERM_DIST), 0) + 1 & " DB @" & vbCrLf & BERM_DIST & "m C/C"
                                            Mtext1_UP.TextHeight = 2.5
                                            Mtext1_UP.Rotation = 0
                                            Mtext1_UP.Location = New Point3d((Linie2.StartPoint.X + Linie2.EndPoint.X) / 2, (Linie2.StartPoint.Y + Linie2.EndPoint.Y) / 2, z)
                                            Mtext1_UP.Attachment = AttachmentPoint.MiddleCenter
                                            BTrecord.AppendEntity(Mtext1_UP)
                                            Trans1.AddNewlyCreatedDBObject(Mtext1_UP, True)



                                            Dim Mtext2 As New MText
                                            Mtext2.Contents = Chainage_text
                                            Mtext2.TextHeight = 2.5
                                            Mtext2.Rotation = PI / 2
                                            Mtext2.Location = New Point3d(X, y - 33, z)
                                            Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                            BTrecord.AppendEntity(Mtext2)
                                            Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                            is_BERM = False
                                            BERM_Start = 0
                                            Count_berm = 0

                                            BERM_DIST = 0

                                            Start_Hor_line = New Point3d(0, 0, 0)
                                        Else
                                            MsgBox("PLEASE specify the BERMS START")
                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                            Start_Hor_line = New Point3d(0, 0, 0)
                                            BERM_Start = 0
                                            Exit Sub

                                        End If


                                End Select










                                If RadioButton_right_left.Checked = True Then
                                    X = X - Distanta_intre_coloane - Delta
                                Else
                                    X = X + Distanta_intre_coloane + Delta
                                End If

                            End If
                        Else

                            If is_BERM = True Then
                                If Count_berm = 1 Then
                                    If RadioButton_right_left.Checked = True Then
                                        X = X - 10
                                    Else
                                        X = X + 10
                                    End If
                                End If
                                If Count_berm = 0 Then
                                    If RadioButton_right_left.Checked = True Then
                                        X = X - 20
                                    Else
                                        X = X + 20
                                    End If
                                End If
                                Dim Linie1 As New Line(New Point3d(X, y - 7.908, z), New Point3d(X, y + 30, z))
                                BTrecord.AppendEntity(Linie1)
                                Trans1.AddNewlyCreatedDBObject(Linie1, True)

                                Dim Linie2 As New Line(Start_Hor_line, New Point3d(X, y + 27, z))

                                Dim Poly1 As New Polyline
                                Poly1.AddVertexAt(0, New Point2d(Linie2.StartPoint.X, Linie2.StartPoint.Y), 0, 0, 2)


                                Dim Minus As Integer = 1
                                If RadioButton_right_left.Checked = True Then
                                    Minus = -1
                                End If

                                Poly1.AddVertexAt(1, New Point2d(Linie2.StartPoint.X + 5 * Minus, Linie2.StartPoint.Y), 0, 0, 0)
                                Poly1.AddVertexAt(2, New Point2d(Linie2.EndPoint.X - 5 * Minus, Linie2.StartPoint.Y), 0, 2, 0)
                                Poly1.AddVertexAt(3, New Point2d(Linie2.EndPoint.X, Linie2.StartPoint.Y), 0, 0, 0)
                                Poly1.Elevation = z


                                BTrecord.AppendEntity(Poly1)
                                Trans1.AddNewlyCreatedDBObject(Poly1, True)




                                Dim Mtext1_UP As New MText
                                Mtext1_UP.Contents = Round(((CDbl(Replace(Chainage_text, "+", "")) - BERM_Start) / BERM_DIST), 0) & " DB @" & vbCrLf & BERM_DIST & "m C/C"
                                Mtext1_UP.TextHeight = 2.5
                                Mtext1_UP.Rotation = 0
                                Mtext1_UP.Location = New Point3d((Linie2.StartPoint.X + Linie2.EndPoint.X) / 2, (Linie2.StartPoint.Y + Linie2.EndPoint.Y) / 2, z)
                                Mtext1_UP.Attachment = AttachmentPoint.MiddleCenter
                                BTrecord.AppendEntity(Mtext1_UP)
                                Trans1.AddNewlyCreatedDBObject(Mtext1_UP, True)



                                Dim Mtext2 As New MText
                                Mtext2.Contents = Chainage_text
                                Mtext2.TextHeight = 2.5
                                Mtext2.Rotation = PI / 2
                                Mtext2.Location = New Point3d(X, y - 33, z)
                                Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                BTrecord.AppendEntity(Mtext2)
                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)

                                If RadioButton_right_left.Checked = True Then
                                    X = X - Distanta_intre_coloane - Delta
                                Else
                                    X = X + Distanta_intre_coloane + Delta
                                End If

                            End If


                            Dim Mtext1 As New MText
                            Mtext1.Contents = PageNR
                            Mtext1.TextHeight = 20
                            Mtext1.Attachment = AttachmentPoint.MiddleRight
                            Mtext1.Layer = "0"
                            Mtext1.Rotation = 0
                            Mtext1.Location = New Point3d(X, y, z)

                            BTrecord.AppendEntity(Mtext1)
                            Trans1.AddNewlyCreatedDBObject(Mtext1, True)

                            X = X0
                            y = y - 125

                            If is_BERM = True Then
                                Start_Hor_line = New Point3d(X, y + 27, z)

                                Dim Linie1 As New Line(New Point3d(X, y - 7.908, z), New Point3d(X, y + 30, z))
                                BTrecord.AppendEntity(Linie1)
                                Trans1.AddNewlyCreatedDBObject(Linie1, True)

                                Dim Mtext2 As New MText
                                Mtext2.Contents = Chainage_text
                                Mtext2.TextHeight = 2.5
                                Mtext2.Rotation = PI / 2
                                Mtext2.Location = New Point3d(X, y - 33, z)
                                Mtext2.Attachment = AttachmentPoint.MiddleLeft
                                BTrecord.AppendEntity(Mtext2)
                                Trans1.AddNewlyCreatedDBObject(Mtext2, True)


                                Count_berm = 0
                                BERM_Start = CDbl(Replace(Chainage_text, "+", ""))

                                If RadioButton_right_left.Checked = True Then
                                    X = X - Distanta_intre_coloane - Delta
                                Else
                                    X = X + Distanta_intre_coloane + Delta
                                End If

                            End If






                        End If



                    Next

                        Trans1.Commit()

                    End Using
                End Using



                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception

                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            afiseaza_butoanele_pentru_forms(Me, Colectie1)


    End Sub




    Private Sub Button_read_from_Excel_Click(sender As Object, e As EventArgs) Handles Button_read_from_Excel.Click
        Survey_band_form_Load(sender, e)
        If IsNumeric(TextBox_row_start.Text) = False Then
            MsgBox("No start row")
            Exit Sub
        End If
        If IsNumeric(TextBox_row_end.Text) = False Then
            MsgBox("No end row")
            Exit Sub
        End If
        If CDbl(TextBox_row_start.Text) < 1 Then
            MsgBox("Start row can't be smaller than 1")
            Exit Sub
        End If
        If CDbl(TextBox_row_end.Text) < 1 Then
            MsgBox("End row can't be smaller than 1")
            Exit Sub
        End If
        If CDbl(TextBox_row_end.Text) < CDbl(TextBox_row_start.Text) Then
            MsgBox("End row can't be smaller than start row")
            Exit Sub
        End If
        Try

            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim Start1 As Integer = CInt(TextBox_row_start.Text)
            Dim End1 As Integer = CInt(TextBox_row_end.Text)

            Dim index_dt As Double = 0

            For i = Start1 To End1
                If Not Replace(W1.Range("A" & i).Value, " ", "") = "" And Not Replace(W1.Range("B" & i).Value, " ", "") = "" And Not Replace(W1.Range("C" & i).Value, " ", "") = "" Then
                    If IsNumeric(Replace(W1.Range("A" & i).Value, "+", "")) = True Then
                        Data_table_survey.Rows.Add()
                        Data_table_survey.Rows(index_dt).Item("CHAINAGE") = CDbl(Replace(W1.Range("A" & i).Value, "+", ""))
                        Data_table_survey.Rows(index_dt).Item("DESCRIPTION") = W1.Range("B" & i).Value
                        Dim CROSSING_TYPE As String = W1.Range("C" & i).Value
                        Data_table_survey.Rows(index_dt).Item("CROSSING_TYPE") = CROSSING_TYPE.ToUpper

                        If Not CROSSING_TYPE.ToUpper = "MATCHLINE" Then
                            If IsNumeric(W1.Range("D" & i).Value) = True Then
                                Data_table_survey.Rows(index_dt).Item("PAGE") = CInt(W1.Range("D" & i).Value)
                            Else
                                MsgBox("PLEASE specify the page number on column D row " & i)
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If
                        End If

                        If IsNumeric(W1.Range("E" & i).Value) = True Then
                            Data_table_survey.Rows(index_dt).Item("BERMDIST") = W1.Range("E" & i).Value
                        End If

                        index_dt = index_dt + 1
                    End If




                Else
                    If Not W1.Range("B" & i).Value.ToUpper = "MATCHLINE" Then W1.Range("A" & i).Interior.ColorIndex = 5
                End If
            Next

            afiseaza_butoanele_pentru_forms(Me, Colectie1)

        Catch ex As Exception
            MsgBox(ex.Message)
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        End Try

    End Sub



End Class