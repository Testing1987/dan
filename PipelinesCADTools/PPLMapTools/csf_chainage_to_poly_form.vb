Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class csf_chainage_to_poly_form
    Dim Data_Table_cl As System.Data.DataTable
    Dim Data_Table_new_sta As System.Data.DataTable
    Dim Data_Table_mtext As System.Data.DataTable
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub csf_chainage_to_poly_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_Table_cl = New System.Data.DataTable
        Data_Table_cl.Columns.Add("X", GetType(Double))
        Data_Table_cl.Columns.Add("Y", GetType(Double))
        Data_Table_cl.Columns.Add("Z", GetType(Double))
        Data_Table_cl.Columns.Add("CSF", GetType(Double))
        Data_Table_cl.Columns.Add("STA", GetType(Double))

        Data_Table_new_sta = New System.Data.DataTable
        Data_Table_new_sta.Columns.Add("NEWSTA", GetType(Double))

        Data_Table_mtext = New System.Data.DataTable
        Data_Table_mtext.Columns.Add("MTEXT", GetType(String))
        Data_Table_mtext.Columns.Add("NEWSTA", GetType(Double))
    End Sub
    Dim POLY3D_EXCEL As Polyline3d
    Private Sub Button_load_CL_Click(sender As Object, e As EventArgs) Handles Button_load_CL.Click
        Try


            If TextBox_x.Text = "" Then
                With TextBox_x
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify X column")
                Exit Sub
            End If

            If TextBox_y.Text = "" Then
                With TextBox_y
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify y column")
                Exit Sub
            End If

            If TextBox_Z.Text = "" Then
                With TextBox_Z
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify z column")
                Exit Sub
            End If

            If TextBox_CSF.Text = "" Then
                With TextBox_CSF
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify CSF column")
                Exit Sub
            End If

            If TextBox_sta.Text = "" Then
                With TextBox_sta
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify STA column")
                Exit Sub
            End If

            If Val(TextBox_start_row.Text) < 1 Or IsNumeric(TextBox_start_row.Text) = False Then
                With TextBox_start_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If
            If Val(TextBox_end_row.Text) < 1 Or IsNumeric(TextBox_end_row.Text) = False Then
                With TextBox_end_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")

                Exit Sub
            End If

            If Val(TextBox_end_row.Text) < Val(TextBox_start_row.Text) Then
                With TextBox_end_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row smaller than start row")

                Exit Sub
            End If
            Dim Start1 As Double = Round(CDbl(TextBox_start_row.Text), 0)
            Dim End1 As Double = Round(CDbl(TextBox_end_row.Text), 0)

            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Data_Table_cl.Rows.Clear()
            Dim Index_cl As Double = 0
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

            Using LOCK1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                    Dim BlockTable As BlockTable = Trans1.GetObject(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.BlockTableId, OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    POLY3D_EXCEL = New Polyline3d
                    BTrecord.AppendEntity(POLY3D_EXCEL)
                    Trans1.AddNewlyCreatedDBObject(POLY3D_EXCEL, True)

                    For i = Start1 To End1
                        Dim X_string As String = W1.Range(TextBox_x.Text & i).Value
                        Dim Y_string As String = W1.Range(TextBox_y.Text & i).Value
                        Dim Z_string As String = W1.Range(TextBox_Z.Text & i).Value
                        Dim CSF_string As String = W1.Range(TextBox_CSF.Text & i).Value
                        Dim STA_string As String = W1.Range(TextBox_sta.Text & i).Value


                        If IsNumeric(X_string) = True And IsNumeric(Y_string) = True And IsNumeric(Z_string) = True And IsNumeric(CSF_string) = True And IsNumeric(STA_string) = True Then
                            Dim X As Double = CDbl(X_string)
                            Dim Y As Double = CDbl(Y_string)
                            Dim Z As Double = CDbl(Z_string)


                            Dim CSF As Double = CDbl(CSF_string)
                            Dim STA As Double = CDbl(STA_string)
                            Data_Table_cl.Rows.Add()
                            Data_Table_cl.Rows(Index_cl).Item("X") = X
                            Data_Table_cl.Rows(Index_cl).Item("Y") = Y
                            Data_Table_cl.Rows(Index_cl).Item("Z") = Z
                            Data_Table_cl.Rows(Index_cl).Item("CSF") = CSF
                            Data_Table_cl.Rows(Index_cl).Item("STA") = STA
                            POLY3D_EXCEL.AppendVertex(New PolylineVertex3d(New Point3d(X, Y, Z)))

                            Index_cl = Index_cl + 1
                        End If





                    Next

                    Trans1.Commit()
                End Using
            End Using
            afiseaza_butoanele_pentru_forms(Me, Colectie1)


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub


    Private Sub Button_LOAD_NEW_STA_Click(sender As Object, e As EventArgs) Handles Button_LOAD_NEW_STA.Click
        Try


            If TextBox_new_sta.Text = "" Then
                With TextBox_new_sta
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify NEW stations column")
                Exit Sub
            End If


            If Val(TextBox_start_row.Text) < 1 Or IsNumeric(TextBox_start_row.Text) = False Then
                With TextBox_start_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If
            If Val(TextBox_end_row.Text) < 1 Or IsNumeric(TextBox_end_row.Text) = False Then
                With TextBox_end_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")

                Exit Sub
            End If

            If Val(TextBox_end_row.Text) < Val(TextBox_start_row.Text) Then
                With TextBox_end_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row smaller than start row")

                Exit Sub
            End If
            Dim Start1 As Double = Round(CDbl(TextBox_start_row.Text), 0)
            Dim End1 As Double = Round(CDbl(TextBox_end_row.Text), 0)

            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Data_Table_new_sta.Rows.Clear()
            Dim Index_NS As Double = 0
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()


            For i = Start1 To End1
                Dim STA_string As String = W1.Range(TextBox_new_sta.Text & i).Value
                If IsNumeric(STA_string) = True Then
                    Dim STA As Double = CDbl(STA_string)
                    Data_Table_new_sta.Rows.Add()
                    Data_Table_new_sta.Rows(Index_NS).Item("NEWSTA") = STA
                    Index_NS = Index_NS + 1
                End If





            Next


            afiseaza_butoanele_pentru_forms(Me, Colectie1)


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_place_on_poly_Click(sender As Object, e As EventArgs) Handles Button_place_on_poly.Click

        Try

            If IsNothing(POLY3D_EXCEL) = True Then
                Exit Sub
            End If
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Editor1 = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()




            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim BlockTable As BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


               
                        ascunde_butoanele_pentru_forms(Me, Colectie1)

                        If Data_Table_new_sta.Rows.Count > 0 And Data_Table_cl.Rows.Count > 0 Then
                            For i = 0 To Data_Table_new_sta.Rows.Count - 1
                                If IsDBNull(Data_Table_new_sta.Rows(i).Item("NEWSTA")) = False Then
                                    Dim NewSta As Double = Data_Table_new_sta.Rows(i).Item("NEWSTA")
                                    Dim Diferenta As Double = 100
                                    Dim Index_close As Double = -1
                                    For j = 0 To Data_Table_cl.Rows.Count - 1
                                        If IsDBNull(Data_Table_cl.Rows(j).Item("STA")) = False Then
                                            Dim Sta As Double = Data_Table_cl.Rows(j).Item("STA")
                                            If Abs(Sta - NewSta) < Abs(Diferenta) Then
                                                Diferenta = NewSta - Sta
                                                Index_close = j
                                            End If
                                        End If
                                    Next
                                    If Not Diferenta = 100 And Not Index_close = -1 Then
                                        If IsDBNull(Data_Table_cl.Rows(Index_close).Item("X")) = False And _
                                            IsDBNull(Data_Table_cl.Rows(Index_close).Item("Y")) = False And _
                                             IsDBNull(Data_Table_cl.Rows(Index_close).Item("Z")) = False And _
                                              IsDBNull(Data_Table_cl.Rows(Index_close).Item("CSF")) = False And _
                                               IsDBNull(Data_Table_cl.Rows(Index_close).Item("STA")) = False Then

                                            Dim X As Double = Data_Table_cl.Rows(Index_close).Item("X")
                                            Dim Y As Double = Data_Table_cl.Rows(Index_close).Item("Y")
                                            Dim Z As Double = Data_Table_cl.Rows(Index_close).Item("Z")
                                            Dim CSF As Double = Data_Table_cl.Rows(Index_close).Item("CSF")
                                            Dim STA As Double = Data_Table_cl.Rows(Index_close).Item("STA")


                                            Dim Real_STA As Double = POLY3D_EXCEL.GetDistAtPoint(New Point3d(X, Y, Z))

                                            Dim Point_at_sta = POLY3D_EXCEL.GetPointAtDist(Real_STA + Diferenta * CSF)
                                            Dim Mleader1 As New MLeader

                                            If IsNothing(Point_at_sta) = False Then
                                                Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_at_sta, Get_chainage_from_double(NewSta, 1), 5, 2.5, 2.5, 5, 10)
                                            End If




                                        End If




                                    End If


                                End If

                            Next
                        End If









                        Trans1.Commit()

                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                End Using ' asta e de la trans1
            End Using


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_add_Mtext_Click(sender As Object, e As EventArgs) Handles Button_add_Mtext.Click
        Try

            If IsNothing(POLY3D_EXCEL) = True Then
                Exit Sub
            End If
            '** added to command list
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Editor1 = ThisDrawing.Editor
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()




            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)



                   

                    Dim BlockTable As BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


                   
                        ascunde_butoanele_pentru_forms(Me, Colectie1)

                        If Data_Table_mtext.Rows.Count > 0 And Data_Table_cl.Rows.Count > 0 Then
                            For i = 0 To Data_Table_mtext.Rows.Count - 1
                                If IsDBNull(Data_Table_mtext.Rows(i).Item("NEWSTA")) = False And IsDBNull(Data_Table_mtext.Rows(i).Item("MTEXT")) = False Then
                                    Dim NewSta As Double = Data_Table_mtext.Rows(i).Item("NEWSTA")
                                    Dim Diferenta As Double = 100
                                    Dim Index_close As Double = -1
                                    For j = 0 To Data_Table_cl.Rows.Count - 1
                                        If IsDBNull(Data_Table_cl.Rows(j).Item("STA")) = False Then
                                            Dim Sta As Double = Data_Table_cl.Rows(j).Item("STA")
                                            If Abs(Sta - NewSta) < Abs(Diferenta) Then
                                                Diferenta = NewSta - Sta
                                                Index_close = j
                                            End If
                                        End If
                                    Next
                                    If Not Diferenta = 100 And Not Index_close = -1 Then
                                        If IsDBNull(Data_Table_cl.Rows(Index_close).Item("X")) = False And _
                                            IsDBNull(Data_Table_cl.Rows(Index_close).Item("Y")) = False And _
                                             IsDBNull(Data_Table_cl.Rows(Index_close).Item("Z")) = False And _
                                              IsDBNull(Data_Table_cl.Rows(Index_close).Item("CSF")) = False And _
                                               IsDBNull(Data_Table_cl.Rows(Index_close).Item("STA")) = False Then

                                            Dim X As Double = Data_Table_cl.Rows(Index_close).Item("X")
                                            Dim Y As Double = Data_Table_cl.Rows(Index_close).Item("Y")
                                            Dim Z As Double = Data_Table_cl.Rows(Index_close).Item("Z")
                                            Dim CSF As Double = Data_Table_cl.Rows(Index_close).Item("CSF")
                                            Dim STA As Double = Data_Table_cl.Rows(Index_close).Item("STA")
                                            Dim Real_STA As Double = POLY3D_EXCEL.GetDistAtPoint(New Point3d(X, Y, Z))

                                            Dim Point_at_sta = POLY3D_EXCEL.GetPointAtDist(Real_STA + Diferenta * CSF)

                                            Dim STRING1 As String = Data_Table_mtext.Rows(i).Item("MTEXT")

                                            Dim Mtext1 As New MText
                                            Mtext1.Contents = STRING1
                                            Mtext1.Location = Point_at_sta
                                            Mtext1.TextHeight = 4
                                            Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                            BTrecord.AppendEntity(Mtext1)
                                            Trans1.AddNewlyCreatedDBObject(Mtext1, True)



                                        End If




                                    End If


                                End If

                            Next
                        End If









                    Trans1.Commit()

                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                End Using ' asta e de la trans1
            End Using


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    Private Sub Button_Load_mtext_Click(sender As Object, e As EventArgs) Handles Button_Load_mtext.Click
        Try


            If TextBox_ADD_TEXT.Text = "" Then
                With TextBox_new_sta
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify Text column")
                Exit Sub
            End If

            If TextBox_new_sta.Text = "" Then
                With TextBox_new_sta
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify NEW stations column")
                Exit Sub
            End If

            If Val(TextBox_start_row.Text) < 1 Or IsNumeric(TextBox_start_row.Text) = False Then
                With TextBox_start_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If
            If Val(TextBox_end_row.Text) < 1 Or IsNumeric(TextBox_end_row.Text) = False Then
                With TextBox_end_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")

                Exit Sub
            End If

            If Val(TextBox_end_row.Text) < Val(TextBox_start_row.Text) Then
                With TextBox_end_row
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row smaller than start row")

                Exit Sub
            End If
            Dim Start1 As Double = Round(CDbl(TextBox_start_row.Text), 0)
            Dim End1 As Double = Round(CDbl(TextBox_end_row.Text), 0)

            ascunde_butoanele_pentru_forms(Me, Colectie1)

            Data_Table_mtext.Rows.Clear()
            Dim Index_NS As Double = 0
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()


            For i = Start1 To End1
                Dim Mtextstring_string As String = W1.Range(TextBox_ADD_TEXT.Text & i).Value
                Dim STA_string As String = W1.Range(TextBox_new_sta.Text & i).Value

                If Not Replace(Mtextstring_string, " ", "") = "" And IsNumeric(STA_string) = True Then
                    Dim STA As Double = CDbl(STA_string)
                    Data_Table_mtext.Rows.Add()
                    Data_Table_mtext.Rows(Index_NS).Item("MTEXT") = Mtextstring_string
                    Data_Table_mtext.Rows(Index_NS).Item("NEWSTA") = STA
                    Index_NS = Index_NS + 1
                End If






            Next


            afiseaza_butoanele_pentru_forms(Me, Colectie1)


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub
End Class