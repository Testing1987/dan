
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class ADD_REV_FORM
    Dim Freeze_operations As Boolean = False
    Dim Data_table1 As System.Data.DataTable
    Dim Data_table2 As System.Data.DataTable
    Private Sub Button_SAVE_AS_Click(sender As Object, e As EventArgs) Handles Button_SAVE_AS.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim Atr_name As String = "Atr_name"
                Dim Atr_value As String = "Atr_value"

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Get_active_worksheet_from_Excel()
                Data_table1 = New System.Data.DataTable
                Data_table1.Columns.Add(Atr_name, GetType(String))
                Data_table1.Columns.Add(Atr_value, GetType(String))

                Data_table2 = New System.Data.DataTable
                Data_table2.Columns.Add(Atr_name, GetType(String))
                Data_table2.Columns.Add(Atr_value, GetType(String))

                Dim Index1 As Integer = 0
                Dim Start1 As Integer = 0
                Dim End1 As Integer = 0

                If Not TextBox_BL_NAME1.Text = "" Then
                    If IsNumeric(TextBox_START1.Text) = True Then
                        Start1 = TextBox_START1.Text
                    End If
                    If IsNumeric(TextBox_END1.Text) = True Then
                        End1 = TextBox_END1.Text
                    End If
                    If Start1 > 0 And End1 > 0 And Start1 <= End1 Then
                        For i = Start1 To End1
                            Data_table1.Rows.Add()
                            Data_table1.Rows(Index1).Item(Atr_name) = W1.Range(TextBox_NAME_column1.Text & i)
                            Data_table1.Rows(Index1).Item(Atr_value) = W1.Range(TextBox_VALUE_column1.Text & i)
                            Index1 = Index1 + 1
                        Next
                    End If
                End If

                Dim Index2 As Integer = 0
                Dim Start2 As Integer = 0
                Dim End2 As Integer = 0
                If Not TextBox_Block_name2.Text = "" Then
                    If IsNumeric(TextBox_START2.Text) = True Then
                        Start2 = TextBox_START2.Text
                    End If
                    If IsNumeric(TextBox_END2.Text) = True Then
                        End2 = TextBox_END2.Text
                    End If
                    If Start2 > 0 And End2 > 0 And Start2 <= End2 Then
                        For i = Start2 To End2
                            Data_table2.Rows.Add()
                            Data_table2.Rows(Index2).Item(Atr_name) = W1.Range(TextBox_NAME_column2.Text & i)
                            Data_table2.Rows(Index2).Item(Atr_value) = W1.Range(TextBox_VALUE_column2.Text & i)
                            Index2 = Index2 + 1
                        Next
                    End If
                End If

                If System.IO.Directory.Exists(TextBox_INPUT_FOLDER.Text) = True And System.IO.Directory.Exists(TextBox_INPUT_FOLDER.Text) = True Then
                    Dim Colectie_nume_fisiere As New Specialized.StringCollection

                    For Each File1 As String In My.Computer.FileSystem.GetFiles(TextBox_INPUT_FOLDER.Text, FileIO.SearchOption.SearchTopLevelOnly, "*.dwg")
                        Colectie_nume_fisiere.Add(File1)
                    Next

                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    For j = 0 To Colectie_nume_fisiere.Count - 1
                        Dim Fisier As String = Colectie_nume_fisiere(j)
                        Dim Database1 As New Database(False, True)
                        Database1.ReadDwgFile(Fisier, IO.FileShare.ReadWrite, False, Nothing)
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                            Dim BlockTable1 As BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead)
                            Dim Paperspace1 As BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                        End Using
                    Next
                End If



            Catch ex As System.Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub


    'Dim SS_name As String = System.IO.Path.GetFileNameWithoutExtension()
    'Dim SS_descr As String = System.IO.Path.GetFileNameWithoutExtension()
    'If System.IO.File.Exists(TextBox_xref_model_space.Text) = True Then
End Class