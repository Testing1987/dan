Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Table_to_excel_form
    Dim Colectie1 As New Specialized.StringCollection

    Private Sub Button_Excel_TO_Existing_Table_Click(sender As Object, e As EventArgs) Handles Button_Excel_TO_Existing_Table.Click
        If IsNumeric(TextBox_Row_start.Text) = False Then
            MsgBox("No start row")
            Exit Sub
        End If
        If IsNumeric(TextBox_Row_End.Text) = False Then
            MsgBox("No end row")
            Exit Sub
        End If
        If CDbl(TextBox_Row_start.Text) < 1 Then
            MsgBox("Start row can't be smaller than 1")
            Exit Sub
        End If
        If CDbl(TextBox_Row_End.Text) < 1 Then
            MsgBox("End row can't be smaller than 1")
            Exit Sub
        End If
        If CDbl(TextBox_Row_End.Text) < CDbl(TextBox_Row_start.Text) Then
            MsgBox("End row can't be smaller than start row")
            Exit Sub
        End If

        If TextBox_COLUMNS_FROM.Text = "" Or Len(TextBox_COLUMNS_FROM.Text) > 1 Then
            MsgBox("Specify the start column")
            Exit Sub
        End If
        If TextBox_COLUMN_TO.Text = "" Or Len(TextBox_COLUMN_TO.Text) > 1 Then
            MsgBox("Specify the end column")
            Exit Sub
        End If

        If Asc(TextBox_COLUMNS_FROM.Text) > Asc(TextBox_COLUMN_TO.Text) Then
            Dim tEMP As String = TextBox_COLUMNS_FROM.Text.ToUpper
            TextBox_COLUMNS_FROM.Text = TextBox_COLUMN_TO.Text.ToUpper
            TextBox_COLUMN_TO.Text = tEMP
        End If

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
        Dim Start1 As Integer = CInt(TextBox_Row_start.Text)
        Dim End1 As Integer = CInt(TextBox_Row_End.Text)
        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select an existing autocad table:")
                    Object_Prompt.SetRejectMessage(vbLf & "You did not select a table")
                    Object_Prompt.AddAllowedClass(GetType(Table), True)


                    Rezultat1 = ThisDrawing.Editor.GetEntity(Object_Prompt)


                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    Dim Table1 As Table = Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForWrite)

                    Dim Index_row_table_autocad As Integer = 0
                    Dim Index_column_table_autocad As Integer = 0
                    Dim Column_excel_from As String = TextBox_COLUMNS_FROM.Text.ToUpper
                    Dim Column_excel_to As String = TextBox_COLUMN_TO.Text.ToUpper

                    If Len(Column_excel_from) = 1 And Len(Column_excel_to) = 1 Then
                        For i = Start1 To End1
                            For j = Asc(Column_excel_from) To Asc(Column_excel_to)
                                Dim valoare As String = W1.Range(Chr(j) & i).Text

                                Dim Nr_rows_tabla As Integer = Table1.Rows.Count
                                If Index_row_table_autocad = Nr_rows_tabla And Not Index_row_table_autocad = 0 Then
                                    Table1.InsertRows(Nr_rows_tabla, Table1.Rows(Nr_rows_tabla - 1).Height, 1)
                                    Table1.Rows(Index_row_table_autocad).TextHeight = Table1.Rows(Index_row_table_autocad - 1).TextHeight
                                    Table1.Rows(Index_row_table_autocad).TextStyleId = Table1.Rows(Index_row_table_autocad - 1).TextStyleId
                                    Table1.Rows(Index_row_table_autocad).Alignment = Table1.Rows(Index_row_table_autocad - 1).Alignment
                                End If

                                If IsNothing(valoare) = False Then
                                    Table1.Rows(Index_row_table_autocad).Item(Index_row_table_autocad, Index_column_table_autocad).Value = valoare
                                Else
                                    Table1.Rows(Index_row_table_autocad).Item(Index_row_table_autocad, Index_column_table_autocad).Value = ""
                                End If
                                Index_column_table_autocad = Index_column_table_autocad + 1
                            Next
                            Index_row_table_autocad = Index_row_table_autocad + 1
                            Index_column_table_autocad = 0
                        Next
                    End If




                    Trans1.Commit()


                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_table_to_excel_Click(sender As Object, e As EventArgs) Handles Button_table_to_excel.Click

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
       
        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptEntityResult
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select an existing autocad table:")
                    Object_Prompt.SetRejectMessage(vbLf & "You did not select a table")
                    Object_Prompt.AddAllowedClass(GetType(Table), True)


                    Rezultat1 = ThisDrawing.Editor.GetEntity(Object_Prompt)


                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If

                    Dim Table1 As Table = Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForRead)
                    Dim Column_excel As String = "A"
                    For i = 0 To Table1.Rows.Count - 1
                        For j = 0 To Table1.Columns.Count - 1
                            W1.Range(Column_excel & (i + 1)).Value = Table1.Rows(i).Item(i, j).Value
                            If Len(Column_excel) = 1 Then
                                If Not Column_excel = "Z" Then
                                    Column_excel = Chr(Asc(Column_excel) + 1)
                                Else
                                    Column_excel = "AA"
                                End If
                            End If

                            If Len(Column_excel) = 2 Then
                                If Not Strings.Right(Column_excel, 1) = "Z" Then
                                    Column_excel = Strings.Left(Column_excel, 1) & Chr(Asc(Strings.Right(Column_excel, 1)) + 1)
                                Else
                                    Column_excel = Chr(Asc(Strings.Left(Column_excel, 1)) + 1) & "A"
                                End If
                            End If
                        Next
                        Column_excel = "A"
                    Next



                    Trans1.Commit()


                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_create_new_table_Click(sender As Object, e As EventArgs) Handles Button_create_new_table.Click
        If IsNumeric(TextBox_Row_start.Text) = False Then
            MsgBox("No start row")
            Exit Sub
        End If
        If IsNumeric(TextBox_Row_End.Text) = False Then
            MsgBox("No end row")
            Exit Sub
        End If
        If CDbl(TextBox_Row_start.Text) < 1 Then
            MsgBox("Start row can't be smaller than 1")
            Exit Sub
        End If
        If CDbl(TextBox_Row_End.Text) < 1 Then
            MsgBox("End row can't be smaller than 1")
            Exit Sub
        End If
        If CDbl(TextBox_Row_End.Text) < CDbl(TextBox_Row_start.Text) Then
            MsgBox("End row can't be smaller than start row")
            Exit Sub
        End If

        If TextBox_COLUMNS_FROM.Text = "" Or Len(TextBox_COLUMNS_FROM.Text) > 1 Then
            MsgBox("Specify the start column")
            Exit Sub
        End If
        If TextBox_COLUMN_TO.Text = "" Or Len(TextBox_COLUMN_TO.Text) > 1 Then
            MsgBox("Specify the end column")
            Exit Sub
        End If
        If Asc(TextBox_COLUMNS_FROM.Text) > Asc(TextBox_COLUMN_TO.Text) Then
            Dim tEMP As String = TextBox_COLUMNS_FROM.Text.ToUpper
            TextBox_COLUMNS_FROM.Text = TextBox_COLUMN_TO.Text.ToUpper
            TextBox_COLUMN_TO.Text = tEMP
        End If

        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
        Dim Start1 As Integer = CInt(TextBox_Row_start.Text)
        Dim End1 As Integer = CInt(TextBox_Row_End.Text)
        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Dim Rezultat_point As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                   
                    Dim curent_ucs_matrix As Matrix3d = ThisDrawing.Editor.CurrentUserCoordinateSystem

                    Rezultat_point = ThisDrawing.Editor.GetPoint(vbLf & "Specify insertion point")


                    If Rezultat_point.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If



                    Dim Index_row_table_autocad As Integer = 0
                    Dim Index_column_table_autocad As Integer = 0
                    Dim Column_excel_from As String = TextBox_COLUMNS_FROM.Text.ToUpper
                    Dim Column_excel_to As String = TextBox_COLUMN_TO.Text.ToUpper

                    If Len(Column_excel_from) = 1 And Len(Column_excel_to) = 1 Then

                        Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                        Dim Exista As Boolean = False
                        Dim Text_style_romans As TextStyleTableRecord

                        For Each TextStyle_id As ObjectId In Text_style_table
                            Dim TextStyle As TextStyleTableRecord = Trans1.GetObject(TextStyle_id, OpenMode.ForRead)
                            With TextStyle
                                If .FileName = "romans.shx" And .XScale = 1.0 And .ObliquingAngle = 0 And Not .Name.ToUpper = "STANDARD" Then
                                    Exista = True
                                    Text_style_romans = TextStyle
                                    Exit For
                                End If
                            End With
                        Next


                        If Exista = False Then
                            Text_style_table.UpgradeOpen()
                            Text_style_romans = New TextStyleTableRecord
                            Text_style_romans.Name = "ROMANS"

                            Text_style_romans.TextSize = 0
                            Text_style_romans.ObliquingAngle = 0
                            Text_style_romans.FileName = "romans.shx"
                            Text_style_romans.XScale = 1.0
                            Text_style_table.Add(Text_style_romans)
                            Trans1.AddNewlyCreatedDBObject(Text_style_romans, True)

                        End If

                        Dim Table1 As New Table
                        Dim Nr_columns As Integer = Asc(Column_excel_to) - Asc(Column_excel_from) + 1
                        Dim Nr_rows As Integer = End1 - Start1 + 1
                        Table1.NumRows = Nr_rows
                        Table1.NumColumns = Nr_columns
                        Table1.SetRowHeight(7.5)
                        Table1.SetColumnWidth(35)
                        Table1.Position = Rezultat_point.Value.TransformBy(curent_ucs_matrix)

                        For i = 0 To Nr_rows - 1
                            Table1.Rows(i).TextStyleId = Text_style_romans.ObjectId
                            Table1.Rows(i).TextHeight = 2.5
                            Table1.Rows(i).Alignment = CellAlignment.MiddleCenter
                        Next



                        For i = Start1 To End1
                            For j = Asc(Column_excel_from) To Asc(Column_excel_to)
                                Dim valoare As String = W1.Range(Chr(j) & i).Value

                                Dim Nr_rows_tabla As Integer = Table1.Rows.Count
                                If Index_row_table_autocad = Nr_rows_tabla And Not Index_row_table_autocad = 0 Then
                                    Table1.InsertRows(Nr_rows_tabla, Table1.Rows(Nr_rows_tabla - 1).Height, 1)
                                    Table1.Rows(Index_row_table_autocad).TextHeight = Table1.Rows(Index_row_table_autocad - 1).TextHeight
                                    Table1.Rows(Index_row_table_autocad).TextStyleId = Table1.Rows(Index_row_table_autocad - 1).TextStyleId
                                    Table1.Rows(Index_row_table_autocad).Alignment = Table1.Rows(Index_row_table_autocad - 1).Alignment
                                End If

                                If IsNothing(valoare) = False Then
                                    Table1.Rows(Index_row_table_autocad).Item(Index_row_table_autocad, Index_column_table_autocad).Value = valoare
                                Else
                                    Table1.Rows(Index_row_table_autocad).Item(Index_row_table_autocad, Index_column_table_autocad).Value = ""
                                End If
                                Index_column_table_autocad = Index_column_table_autocad + 1
                            Next
                            Index_row_table_autocad = Index_row_table_autocad + 1
                            Index_column_table_autocad = 0
                        Next




                        BTrecord.AppendEntity(Table1)
                        Trans1.AddNewlyCreatedDBObject(Table1, True)

                        Trans1.Commit()

                    End If
                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub
End Class