Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Data.OleDb

Public Class Workspace_band_form
    Dim Colectie1 As New Specialized.StringCollection
    Dim PolyCL As Polyline
    Dim Data_table_matchlines As System.Data.DataTable
    Dim Index_data_table As Integer = 0

    Private Sub Workspace_band_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_table_matchlines = New System.Data.DataTable
        Data_table_matchlines.Columns.Add("MATCHLINE", GetType(Double))
    End Sub
    Private Sub Button_Load_CL_Click(sender As Object, e As EventArgs) Handles Button_Load_CL.Click


        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the centerline:")
                    Object_Prompt.SetRejectMessage(vbLf & "You did not select a polyline")
                    Object_Prompt.AddAllowedClass(GetType(Polyline), True)


                    Rezultat1 = ThisDrawing.Editor.GetEntity(Object_Prompt)


                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    PolyCL = Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForRead)


                    Trans1.Commit()


                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        afiseaza_butoanele_pentru_forms(Me, Colectie1)

    End Sub

    Private Sub Button_load_matchlines_Click(sender As Object, e As EventArgs) Handles Button_load_matchlines.Click
        If IsNothing(PolyCL) = False Then
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Try
                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                        Dim Curent_UCS As Matrix3d = ThisDrawing.Editor.CurrentUserCoordinateSystem
                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Matchline polylines"

                        Object_Prompt.SingleOnly = False
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Object_Prompt)


                        If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                            Exit Sub
                        End If


                        For i = 0 To Rezultat1.Value.Count - 1
                            Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf ent1 Is Polyline Then
                                Dim Poly1 As Polyline = ent1
                                If Poly1.Closed = True Then
                                    Dim Col_int As New Point3dCollection
                                    Poly1.IntersectWith(PolyCL, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero)
                                    If Col_int.Count > 0 Then
                                        For j = 0 To Col_int.Count - 1
                                            Dim Chainage_at_pt As Double = Round(PolyCL.GetDistAtPoint(Col_int(j).TransformBy(Curent_UCS)), 1)

                                            If Chainage_at_pt > 10 Then
                                                Dim Exista_deja As Boolean = False
                                                For k = 0 To Data_table_matchlines.Rows.Count - 1
                                                    If Data_table_matchlines.Rows(k).Item("MATCHLINE") = Chainage_at_pt Then
                                                        Exista_deja = True
                                                    End If
                                                Next
                                                If Exista_deja = False Then
                                                    Data_table_matchlines.Rows.Add()
                                                    Data_table_matchlines.Rows(Index_data_table).Item("MATCHLINE") = Chainage_at_pt
                                                    Index_data_table = Index_data_table + 1
                                                End If
                                            End If

                                        Next

                                    End If
                                End If
                            End If

                        Next


                        Data_table_matchlines = Sort_data_table(Data_table_matchlines, "MATCHLINE")

                        ListBox_matchlines.Items.Clear()

                        For i = 0 To Data_table_matchlines.Rows.Count - 1
                            ListBox_matchlines.Items.Add(Data_table_matchlines.Rows(i).Item("MATCHLINE"))
                        Next


                        Trans1.Commit()


                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        End If

    End Sub


    Private Sub Button_DRAW_WS_Click(sender As Object, e As EventArgs) Handles Button_DRAW_WS.Click, Button2.Click, Button1.Click, Button3.Click, Button4.Click, Button6.Click, Button5.Click
        If IsNothing(PolyCL) = False Then
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Try
                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                        Dim Curent_UCS As Matrix3d = ThisDrawing.Editor.CurrentUserCoordinateSystem
                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                        Dim start1 As Integer = CInt(TextBox_rowStart.Text)
                        Dim end1 As Integer = CInt(TextBox_rowEnd.Text)
                        For i = start1 To end1
                            Dim Chainage_start As Double = W1.Range(TextBox_COLUMN_START.Text.ToUpper & i).Value
                            Dim Chainage_end As Double = W1.Range(TextBox_COLUMN_END.Text.ToUpper & i).Value
                            Dim Width As Double = W1.Range(TextBox_width_column.Text.ToUpper & i).Value
                            Dim Off_left As Double = W1.Range(TextBox_Left_offset.Text.ToUpper & i).Value
                            Dim Off_right As Double = W1.Range(TextBox_Rght_offset.Text.ToUpper & i).Value

                            If Chainage_start <= PolyCL.Length And Chainage_end <= PolyCL.Length Then
                                Dim Point_on_CL_start As New Point3d
                                Point_on_CL_start = PolyCL.GetPointAtDist(Chainage_start).TransformBy(Curent_UCS)
                                Dim Point_on_CL_end As New Point3d
                                Point_on_CL_end = PolyCL.GetPointAtDist(Chainage_end).TransformBy(Curent_UCS)
                                Dim Index_first As Integer = Ceiling(PolyCL.GetParameterAtPoint(Point_on_CL_start))
                                Dim Index_last As Integer = Ceiling(PolyCL.GetParameterAtPoint(Point_on_CL_end))

                                Dim Poly_temp As New Polyline
                                Poly_temp.AddVertexAt(0, New Point2d(Point_on_CL_start.X, Point_on_CL_start.Y), 0, 0, 0)





                            End If
                        Next



                        Trans1.Commit()


                    End Using
                End Using

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click, Label22.Click, Label26.Click, Label40.Click, Label27.Click, Label20.Click

    End Sub

    Private Sub Panel_layers_Paint(sender As Object, e As Windows.Forms.PaintEventArgs) Handles Panel_layers.Paint

    End Sub

    Function GetData() As System.Data.DataSet

        Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0; Ole DB Services=-4; Data Source=D:\Tecnical Stu" & _
       "dy\Complete_Code\Ch08\data\NorthWind.mdb"
        Dim dbConnection As System.Data.IDbConnection = New System.Data.OleDb.OleDbConnection(connectionString)

        Dim queryString As String = "SELECT [Employees].* FROM [Employees]"
        Dim dbCommand As System.Data.IDbCommand = New System.Data.OleDb.OleDbCommand
        dbCommand.CommandText = queryString
        dbCommand.Connection = dbConnection

        Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter
        dataAdapter.SelectCommand = dbCommand
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet
        dataAdapter.Fill(dataSet)

        Return dataSet
    End Function

    Private Sub Button_connect_to_access_DB_Click(sender As Object, e As EventArgs) Handles Button_connect_to_access_DB.Click
        Try

            Dim Table1 As String = "ROW_CONFIG_TABLE"
            Dim query As String = "SELECT * FROM " & Table1
            Dim MDBConnString_ As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\pop70694\Documents\Work Files\2015-09-16 Access database\KM_Workspace.accdb;"
            Dim ds As New DataSet
            Dim cnn As OleDbConnection

            Try
                cnn = New OleDbConnection(MDBConnString_)
                cnn.Open()
                Dim cmd As New OleDbCommand(query, cnn)
                Dim da As New OleDbDataAdapter(cmd)
                da.Fill(ds, Table1)
                cnn.Close()
            Catch ex As OleDb.OleDbException
                MsgBox(ex.Message)
                Exit Sub
            End Try
            Dim t1 As System.Data.DataTable = ds.Tables(Table1)
            MsgBox(t1.Columns(4).ColumnName & " Row 2 = " & t1.Rows(1).Item(4))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class