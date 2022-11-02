Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Commands_class_AL_BAND
    <CommandMethod("CPU3")> _
    Public Sub gaseste_nr_serial_al_disc_C()
        Dim disk As New Management.ManagementObject("Win32_LogicalDisk.DeviceID=""C:""")
        Dim diskPropertyB As Management.PropertyData = disk.Properties("VolumeSerialNumber")
        MsgBox(diskPropertyB.Value.ToString())
    End Sub

    <CommandMethod("PPL_matcoatband")> _
    Public Sub Show_mat_and_coat_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is ALIGNMENT_MATERIAL_AND_COATING_FORM Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New ALIGNMENT_MATERIAL_AND_COATING_FORM
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_PROPBAND")> _
    Public Sub Show_Line_list_band_Form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Custom_User_band_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Custom_User_band_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("PPL_engband")> _
    Public Sub Show_engineering_band_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Engineering_band_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Engineering_band_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("graph_convertor")> _
    Public Sub Show_graph_conv_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Graph_converter Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Graph_converter
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("Point_at_station")> _
    Public Sub point_at_station_us_style()

        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If
            Dim Poly2D As Polyline

            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d




            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1



                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim String1 As Autodesk.AutoCAD.EditorInput.PromptStringOptions
                String1 = New Autodesk.AutoCAD.EditorInput.PromptStringOptions(vbLf & "Specify station:")
                String1.AllowSpaces = True

                Dim Descriptia As Autodesk.AutoCAD.EditorInput.PromptResult = Editor1.GetString(String1)

                If Descriptia.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    Exit Sub
                End If

                Dim Ch_result As String = Descriptia.StringResult
                Ch_result = Replace(Ch_result, "+", "")
                Ch_result = Replace(Ch_result, " ", "")
                If IsNumeric(Ch_result) = False Then
                    MsgBox("Station is not specified correctly")
                    Exit Sub

                End If
                Dim Chainage As Double = CDbl(Ch_result)

                If Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If
                If IsNothing(Poly2D) = False Then
                    Point_on_poly = Poly2D.GetPointAtDist(Chainage)
                End If
                If IsNothing(Poly3D) = False Then
                    Point_on_poly = Poly3D.GetPointAtDist(Chainage)
                End If

                Dim Chainage_string As String = Get_chainage_feet_from_double(Chainage, 2)


                If Chainage_string = "-0+00.00" Then Chainage_string = "0+00.00"

                If IsNothing(Point_on_poly) = False Then
                    Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 50, 2.5, 20, 50, 100)
                End If

                Trans1.Commit()

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("station_at_point")> _
    Public Sub station_at_point_us_style()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True
            Rezultat1 = Editor1.GetSelection(Object_Prompt)

            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly2D As Polyline
            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1


                            Dim Point_zero As New Point3d

                            Point_zero = Poly2D.GetClosestPointTo(Poly2D.StartPoint, Vector3d.ZAxis, False)



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Dim Index_Poly As Integer = 0
                            Poly2D = New Polyline
                            For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In Poly3D
                                Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                        (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                Dim x1 As Double = v3d.Position.X
                                Dim y1 As Double = v3d.Position.Y
                                Dim z1 As Double = v3d.Position.Z
                                Poly2D.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                Index_Poly = Index_Poly + 1
                            Next



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the same polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly2D) = False Then
                    Point_on_poly = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly2D.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D) = False Then
                    Dim Point_on_poly2D As New Point3d
                    Point_on_poly2D = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Dim param1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2D)
                    Distanta_pana_la_xing = Poly3D.GetDistanceAtParameter(param1)
                    Point_on_poly = Poly3D.GetPointAtParameter(param1)
                End If


                Dim Chainage As Double = Distanta_pana_la_xing - Dist_from_start_for_zero




                If Dist_from_start_for_zero + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If

                Dim Chainage_string As String = Get_chainage_feet_from_double(Chainage, 2)
                If Chainage_string = "-0+00.00" Then Chainage_string = "0+00.00"

                Dim Mleader1 As New MLeader

                If IsNothing(Point_on_poly) = False Then
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 50, 2.5, 20, 50, 100)
                End If

                Trans1.Commit()

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("etc_1")> _
    Public Sub station_at_point_pick_elev_transfer_excel()
        If isSECURE() = False Then Exit Sub

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
        W1 = Get_active_worksheet_from_Excel()


        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Line_prompt00 As New Autodesk.AutoCAD.EditorInput.PromptIntegerOptions(vbLf & "Row:")
            Line_prompt00.AllowNegative = False
            Line_prompt00.AllowZero = False
            Line_prompt00.AllowNone = True


            Dim Start_row_result As Autodesk.AutoCAD.EditorInput.PromptIntegerResult = Editor1.GetInteger(Line_prompt00)

            Dim Row1 As Integer = 1

            If Start_row_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Row1 = Start_row_result.Value
            End If

            Dim Poly2D As Polyline
            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1


                            Dim Point_zero As New Point3d

                            Point_zero = Poly2D.GetClosestPointTo(Poly2D.StartPoint, Vector3d.ZAxis, False)



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Dim Index_Poly As Integer = 0
                            Poly2D = New Polyline
                            For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In Poly3D
                                Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                        (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                Dim x1 As Double = v3d.Position.X
                                Dim y1 As Double = v3d.Position.Y
                                Dim z1 As Double = v3d.Position.Z
                                Poly2D.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                Index_Poly = Index_Poly + 1
                            Next



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the same polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly2D) = False Then
                    Point_on_poly = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly2D.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D) = False Then
                    Dim Point_on_poly2D As New Point3d
                    Point_on_poly2D = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Dim param1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2D)
                    Distanta_pana_la_xing = Poly3D.GetDistanceAtParameter(param1)
                    Point_on_poly = Poly3D.GetPointAtParameter(param1)
                End If


                Dim Chainage As Double = Distanta_pana_la_xing - Dist_from_start_for_zero




                If Dist_from_start_for_zero + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If




                Dim Chainage_string As String = Get_chainage_feet_from_double(Chainage, 0)
                If Chainage_string = "-0+00" Then Chainage_string = "0+00"

                Dim Mleader1 As New MLeader

                If IsNothing(Point_on_poly) = False Then
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 2, 1, 1, 1, 8)
                End If



                Dim Elev_prompt00 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Elev:")
                Elev_prompt00.AllowNegative = False
                Elev_prompt00.AllowZero = False
                Elev_prompt00.AllowNone = True
                'Dim elevation_result As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Elev_prompt00)
                If Start_row_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Dim Rezultat_block As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt_block_ref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt_block_ref.MessageForAdding = vbLf & "Select survey block:"

                    Object_Prompt_block_ref.SingleOnly = True

                    Rezultat_block = Editor1.GetSelection(Object_Prompt_block_ref)

                    Try


                        If Rezultat_block.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Dim Block1 As BlockReference = TryCast(Trans1.GetObject(Rezultat_block.Value(0).ObjectId, OpenMode.ForRead), BlockReference)
                            If IsNothing(Block1) = False Then
                                If Block1.AttributeCollection.Count > 0 Then

                                    For Each id As ObjectId In Block1.AttributeCollection
                                        If Not id.IsErased Then
                                            Dim attRef As AttributeReference = DirectCast(Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                            If attRef.Tag.ToUpper = "DESC2" Then
                                                Dim Continut As String = attRef.TextString
                                                W1.Range("c" & Row1).Value2 = Continut
                                            End If
                                            If attRef.Tag.ToUpper = "ELEV2" Then
                                                Dim Continut As String = attRef.TextString
                                                W1.Range("B" & Row1).Value2 = Continut
                                            End If
                                        End If
                                    Next


                                End If
                            End If
                        End If

                    Catch ex As Exception

                    End Try

                    W1.Range("a" & Row1).Value2 = Round(Chainage, 2)

                    Row1 = Row1 + 1
                End If

                Trans1.Commit()
                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub


    <CommandMethod("Clone_and_replace", CommandFlags.UsePickSet)> _
    Public Sub Clone_and_replace_objects()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
        Editor1 = ThisDrawing.Editor
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If
            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If
            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                If IsNothing(Rezultat1) = False Then
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        For i = 0 To Rezultat1.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForWrite)
                            Dim Ent2 As Entity = Ent1.Clone
                            BTrecord.AppendEntity(Ent2)
                            Trans1.AddNewlyCreatedDBObject(Ent2, True)
                            Ent1.Erase()
                        Next
                        Trans1.Commit()
                        Editor1.Regen()
                    End Using
                Else
                    Exit Sub
                End If
            End If
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
