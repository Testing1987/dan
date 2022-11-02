Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class References_from_excel_form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Data_table_Refs As New System.Data.DataTable
    Dim Data_table_Refs_with_sheets As New System.Data.DataTable
    Dim Data_table_REF_USA As New System.Data.DataTable
    Dim Nr_blocks As Integer = 0

    Dim Data_table_Viewport_data As New System.Data.DataTable

    Dim Freeze_operations As Boolean = False

    Private Sub Ref_from_XL_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_table_Refs.Columns.Add("CODE", GetType(String))
        Data_table_Refs.Columns.Add("DESCRIPTION", GetType(String))

        Data_table_Refs_with_sheets.Columns.Add("CODE", GetType(String))
        Data_table_Refs_with_sheets.Columns.Add("DESCRIPTION", GetType(String))
        Data_table_Refs_with_sheets.Columns.Add("SHEET", GetType(Integer))
    End Sub
    Private Sub Button_LOAD_TEXT_FROM_EXCEL_Click(sender As Object, e As EventArgs) Handles Button_LOAD_TEXT_FROM_EXCEL.Click
        Try



            If TextBox_column_XL_DWG.Text = "" Or TextBox_column_XL_descr.Text = "" Then
                MsgBox("Please specify the EXCEL COLUMN!")
                Exit Sub
            End If
            If TextBox_start.Text = "" Then
                MsgBox("Please specify the EXCEL START ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_start.Text) = False Then
                With TextBox_start
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If
            If TextBox_end.Text = "" Then
                MsgBox("Please specify the EXCEL END ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_end.Text) = False Then
                With TextBox_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify END row")

                Exit Sub
            End If

            If Val(TextBox_start.Text) < 1 Then
                With TextBox_start
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Start row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_end.Text) < 1 Then
                With TextBox_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_end.Text) < Val(TextBox_start.Text) Then
                MsgBox("END row smaller than start row")

                Exit Sub
            End If

            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_start.Text)
            Dim end1 As Integer = CInt(TextBox_end.Text)
            Data_table_Refs.Rows.Clear()
            Data_table_Refs_with_sheets.Rows.Clear()

            Dim Index1 As Integer = 0
            Dim Index5 As Integer = 0
            For i = start1 To end1
                Dim Cod As String = W1.Range(TextBox_column_XL_DWG.Text.ToUpper & i).Value
                Dim Description As String = W1.Range(TextBox_column_XL_descr.Text.ToUpper & i).Value
                Dim Pagina_text As String

                If Not W1.Range(TextBox_column_XL_Sheet.Text.ToUpper & i).Text = "" Then
                    Pagina_text = W1.Range(TextBox_column_XL_Sheet.Text.ToUpper & i).Value
                End If

                If Not Cod = "" And Not Description = "" Then
                    Dim Exista As Boolean = False
                    For j = 0 To Data_table_Refs.Rows.Count - 1
                        If Data_table_Refs.Rows(j).Item("CODE").ToString.ToUpper = Cod.ToUpper Then
                            Exista = True
                            Exit For
                        End If
                    Next


                    If Exista = False Then
                        Data_table_Refs.Rows.Add()
                        Data_table_Refs.Rows(Index1).Item("CODE") = Cod
                        Data_table_Refs.Rows(Index1).Item("DESCRIPTION") = Description
                        Index1 = Index1 + 1
                    End If

                    Data_table_Refs_with_sheets.Rows.Add()
                    Data_table_Refs_with_sheets.Rows(Index5).Item("CODE") = Cod
                    Data_table_Refs_with_sheets.Rows(Index5).Item("DESCRIPTION") = Description
                    If Not Pagina_text = "" Then
                        Dim Pagina_integer As Integer = CInt(extrage_numar_din_text_de_la_sfarsitul_textului(Pagina_text))

                        Data_table_Refs_with_sheets.Rows(Index5).Item("SHEET") = Pagina_integer
                    End If


                    Index5 = Index5 + 1



                End If
            Next
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub


    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button_dwg.Click
        Try
            If Data_table_REF_USA.Rows.Count = 0 Then
                MsgBox("You don't have loaded any refs!")
                Exit Sub
            End If

            If TextBox_nr_max_descr_line.Text = "" Or TextBox_nr_max_descr_line.Text = "" Then
                MsgBox("Please specify the number of characters for a single line!")
                Exit Sub
            End If

            If IsNumeric(TextBox_nr_max_dwg_line.Text) = False Or IsNumeric(TextBox_nr_max_descr_line.Text) = False Then
                With TextBox_nr_max_dwg_line
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the number of characters for a single line!")

                Exit Sub
            End If


            If Val(TextBox_nr_max_dwg_line.Text) < 1 Or Val(TextBox_nr_max_descr_line.Text) < 1 Then
                With TextBox_nr_max_dwg_line
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the number of characters for a single line!")

                Exit Sub
            End If

            Dim Nr_max_litere_descr As Integer = CInt(TextBox_nr_max_descr_line.Text)
            Dim Nr_max_litere_dwg As Integer = CInt(TextBox_nr_max_dwg_line.Text)

            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
123:            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select the existing blocks containing DWG numbers:"
                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Dim Data_table_from_blocks As New System.Data.DataTable
                        Data_table_from_blocks.Columns.Add("CODE", GetType(String))
                        Dim index2 As Integer = 0

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is BlockReference Then
                                    Dim Block1 As BlockReference = Ent1
                                    If Block1.AttributeCollection.Count > 0 Then


                                        For Each Atid As ObjectId In Block1.AttributeCollection
                                            Dim Atr1 As AttributeReference = Atid.GetObject(OpenMode.ForRead)


                                            If Atr1.Tag.ToUpper = "ID_NO" Then
                                                Dim String1 As String
                                                If Atr1.IsMTextAttribute = True Then
                                                    String1 = Atr1.MTextAttribute.Text
                                                Else
                                                    String1 = Atr1.TextString
                                                End If
                                                Data_table_from_blocks.Rows.Add()
                                                Data_table_from_blocks.Rows(index2).Item("CODE") = String1
                                                index2 = index2 + 1

                                            End If


                                        Next

                                    End If
                                End If



                            Next

                            If Data_table_from_blocks.Rows.Count > 0 Then
                                Dim Colectie_dwg As New System.Data.DataTable
                                Colectie_dwg.Columns.Add("DWG", GetType(String))
                                Colectie_dwg.Columns.Add("DESCR", GetType(String))

                                Dim Index4 As Integer = 0
                                For i = 0 To Data_table_from_blocks.Rows.Count - 1
                                    Dim Valoare_DWG As String = Data_table_from_blocks.Rows(i).Item("CODE")
                                    For j = 0 To Data_table_Refs.Rows.Count - 1
                                        If Valoare_DWG.ToUpper = Data_table_Refs.Rows(j).Item("CODE").ToString.ToUpper Then
                                            Colectie_dwg.Rows.Add()
                                            Colectie_dwg.Rows(Index4).Item("DWG") = Valoare_DWG
                                            Colectie_dwg.Rows(Index4).Item("DESCR") = Data_table_Refs.Rows(j).Item("DESCRIPTION")
                                            Index4 = Index4 + 1
                                            Exit For

                                        End If


                                    Next
                                Next

                                Colectie_dwg = Sort_data_table(Colectie_dwg, "DWG")

                                If Colectie_dwg.Rows.Count > 0 Then
                                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                                    Object_Prompt2.MessageForAdding = vbLf & "Select the block to be populated:"
                                    Object_Prompt2.SingleOnly = True
                                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                                    If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        If IsNothing(Rezultat2) = False Then
                                            Dim Ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForWrite)
                                            If TypeOf Ent2 Is BlockReference Then
                                                Dim Block2 As BlockReference = Ent2
                                                If Block2.AttributeCollection.Count > 0 Then
                                                    Dim Data_table_Tags As New System.Data.DataTable
                                                    Data_table_Tags.Columns.Add("TAG", GetType(String))
                                                    Data_table_Tags.Columns.Add("X", GetType(Double))
                                                    Data_table_Tags.Columns.Add("Y", GetType(Double))
                                                    Dim Index3 As Integer = 0
                                                    For Each atrID In Block2.AttributeCollection
                                                        Dim Atr2 As AttributeReference = atrID.GetObject(OpenMode.ForWrite)
                                                        Data_table_Tags.Rows.Add()
                                                        Data_table_Tags.Rows(Index3).Item("TAG") = Atr2.Tag
                                                        Data_table_Tags.Rows(Index3).Item("X") = Round(Atr2.Position.X, 0)
                                                        Data_table_Tags.Rows(Index3).Item("Y") = Round(Atr2.Position.Y, 0)
                                                        Index3 = Index3 + 1
                                                        If Atr2.IsMTextAttribute = True Then
                                                            Atr2.MTextAttribute.Contents = ""
                                                        Else
                                                            Atr2.TextString = ""
                                                        End If
                                                    Next

                                                    '" DESC,"  " ASC"

                                                    Data_table_Tags = Sort_data_table_2_columns(Data_table_Tags, "Y", " DESC,", "X", " ASC")
                                                    Dim Colectie_valori As New Specialized.StringCollection
                                                    Dim Colectie_shrink As New IntegerCollection
                                                    For k = 0 To Colectie_dwg.Rows.Count - 1
                                                        If Colectie_valori.Contains(Colectie_dwg.Rows(k).Item("DWG")) = False Then
                                                            Dim Descr2 As String = Colectie_dwg.Rows(k).Item("DESCR")

                                                            If Strings.Len(Descr2) <= Nr_max_litere_descr Then
                                                                Colectie_valori.Add(Colectie_dwg.Rows(k).Item("DWG"))
                                                                If Len(Colectie_dwg.Rows(k).Item("DWG")) > Nr_max_litere_dwg Then
                                                                    Colectie_shrink.Add(1)
                                                                Else
                                                                    Colectie_shrink.Add(0)
                                                                End If
                                                                Colectie_valori.Add(Colectie_dwg.Rows(k).Item("DESCR"))
                                                                Colectie_shrink.Add(0)
                                                            End If


                                                            If Strings.Len(Descr2) > Nr_max_litere_descr Then
                                                                Dim DescrA As String = Strings.Left(Descr2, Nr_max_litere_descr)
                                                                Dim Split_car As Integer
                                                                For l = 1 To Nr_max_litere_descr
                                                                    If Strings.Left(Strings.Right(DescrA, l), 1) = " " Then
                                                                        Split_car = Strings.Len(DescrA) - l
                                                                        Exit For
                                                                    End If
                                                                Next
                                                                Colectie_valori.Add(Colectie_dwg.Rows(k).Item("DWG"))
                                                                If Len(Colectie_dwg.Rows(k).Item("DWG")) > Nr_max_litere_dwg Then
                                                                    Colectie_shrink.Add(1)
                                                                Else
                                                                    Colectie_shrink.Add(0)
                                                                End If
                                                                Colectie_valori.Add(Strings.Left(Descr2, Split_car))
                                                                Colectie_shrink.Add(0)

                                                                Dim DescrB As String = Strings.Right(Descr2, Strings.Len(Descr2) - Split_car - 1)

                                                                If Strings.Len(DescrB) <= Nr_max_litere_descr Then
                                                                    Colectie_valori.Add(" ")
                                                                    Colectie_shrink.Add(0)
                                                                    Colectie_valori.Add(DescrB)
                                                                    Colectie_shrink.Add(0)
                                                                Else
                                                                    Colectie_valori.Add(" ")
                                                                    Colectie_shrink.Add(0)
                                                                    Dim DescrC As String = Strings.Left(DescrB, Nr_max_litere_descr)
                                                                    Dim Split_carB As Integer
                                                                    For l = 1 To Nr_max_litere_descr
                                                                        If Strings.Left(Strings.Right(DescrC, l), 1) = " " Then
                                                                            Split_carB = Strings.Len(DescrC) - l
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                    Colectie_valori.Add(Strings.Left(DescrB, Split_carB))
                                                                    Colectie_shrink.Add(0)

                                                                    Dim DescrB1 As String = Strings.Right(DescrB, Strings.Len(DescrB) - Split_carB - 1)

                                                                    If Strings.Len(DescrB1) <= Nr_max_litere_descr Then
                                                                        Colectie_valori.Add(" ")
                                                                        Colectie_shrink.Add(0)
                                                                        Colectie_valori.Add(DescrB1)
                                                                        Colectie_shrink.Add(0)
                                                                    Else
                                                                        Colectie_valori.Add(" ")
                                                                        Colectie_shrink.Add(0)
                                                                        Dim DescrC1 As String = Strings.Left(DescrB1, Nr_max_litere_descr)
                                                                        Dim Split_carB1 As Integer
                                                                        For l = 1 To Nr_max_litere_descr
                                                                            If Strings.Left(Strings.Right(DescrC1, l), 1) = " " Then
                                                                                Split_carB1 = Strings.Len(DescrC1) - l
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                        Colectie_valori.Add(Strings.Left(DescrB1, Split_carB1))
                                                                        Colectie_shrink.Add(0)
                                                                        Dim DescrB2 As String = Strings.Right(DescrB1, Strings.Len(DescrB1) - Split_carB1 - 1)

                                                                        Colectie_valori.Add(" ")
                                                                        Colectie_shrink.Add(0)
                                                                        Colectie_valori.Add(DescrB2)
                                                                        Colectie_shrink.Add(0)
                                                                    End If
                                                                End If
                                                            End If



                                                        End If

                                                    Next
                                                    Dim Index_colectie As Integer = 0

                                                    For Each atrID In Block2.AttributeCollection

                                                        Dim Atr3 As AttributeReference = atrID.GetObject(OpenMode.ForWrite)
                                                        Dim Tag3 As String = Atr3.Tag



                                                        For m = 0 To Data_table_Tags.Rows.Count - 1
                                                            If Tag3.ToUpper = Data_table_Tags.Rows(m).Item("TAG").ToString.ToUpper Then
                                                                Index_colectie = m
                                                                Exit For
                                                            End If

                                                        Next

                                                        If Index_colectie <= Colectie_shrink.Count - 1 Then
                                                            If Colectie_shrink(Index_colectie) = 1 Then
                                                                Atr3.WidthFactor = 0.9
                                                            Else
                                                                Atr3.WidthFactor = 1
                                                            End If

                                                            If Atr3.IsMTextAttribute = True Then
                                                                Atr3.MTextAttribute.Contents = Colectie_valori(Index_colectie)
                                                            Else
                                                                Atr3.TextString = Colectie_valori(Index_colectie)
                                                            End If
                                                        End If



                                                    Next
                                                    If Colectie_valori.Count > Block2.AttributeCollection.Count Then
                                                        MsgBox("You have more values than block attributes for this sheet!")
                                                    End If

                                                End If
                                            End If
                                        End If
                                    End If

                                End If
                            End If




                            Trans1.Commit()
                        End Using

                    End If
                End If

                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button_read_excel_sheet_Click(sender As Object, e As EventArgs) Handles Button_read_excel_sheet.Click
        Try

            If TextBox_column_XL_Sheet.Text = "" Then
                MsgBox("Please specify the EXCEL COLUMN!")
                Exit Sub
            End If
            If TextBox_SHEET_NO.Text = "" Then
                MsgBox("Please specify the SHEET NUMBER!")
                Exit Sub
            End If

            If IsNumeric(TextBox_SHEET_NO.Text) = False Then
                With TextBox_SHEET_NO
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify SHEET NUMBER!")

                Exit Sub
            End If


            If Val(TextBox_SHEET_NO.Text) < 1 Then
                With TextBox_start
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify SHEET NUMBER!")

                Exit Sub
            End If

            If TextBox_nr_max_descr_line.Text = "" Or TextBox_nr_max_dwg_line.Text = "" Then
                MsgBox("Please specify the number of characters for a single line!")
                Exit Sub
            End If

            If IsNumeric(TextBox_nr_max_descr_line.Text) = False Or IsNumeric(TextBox_nr_max_dwg_line.Text) = False Then
                With TextBox_nr_max_descr_line
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the number of characters for a single line!")

                Exit Sub
            End If


            If Val(TextBox_nr_max_descr_line.Text) < 1 Or Val(TextBox_nr_max_dwg_line.Text) < 1 Then
                With TextBox_nr_max_descr_line
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the number of characters for a single line!")

                Exit Sub
            End If

            If Data_table_Refs_with_sheets.Rows.Count = 0 Then
                MsgBox("You don't have loaded any references!")
                Exit Sub
            End If




            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim Colectie_dwg As New System.Data.DataTable
                    Colectie_dwg.Columns.Add("DWG", GetType(String))
                    Colectie_dwg.Columns.Add("DESCR", GetType(String))

                    Dim Index4 As Integer = 0
                    Dim Pagina1 As Integer = CInt(TextBox_SHEET_NO.Text)
                    TextBox_SHEET_NO.Text = (Pagina1 + 1).ToString

                    Dim Nr_max_litere_descr As Integer = CInt(TextBox_nr_max_descr_line.Text)
                    Dim Nr_max_litere_dwg As Integer = CInt(TextBox_nr_max_dwg_line.Text)
                    For i = 0 To Data_table_Refs_with_sheets.Rows.Count - 1

                        Dim Valoare_pagina As String = Data_table_Refs_with_sheets.Rows(i).Item("SHEET")
                        If Pagina1 = Valoare_pagina Then
                            Colectie_dwg.Rows.Add()
                            Colectie_dwg.Rows(Index4).Item("DWG") = Data_table_Refs_with_sheets.Rows(i).Item("CODE")
                            Colectie_dwg.Rows(Index4).Item("DESCR") = Data_table_Refs_with_sheets.Rows(i).Item("DESCRIPTION")
                            Index4 = Index4 + 1
                        End If
                    Next


                    If Colectie_dwg.Rows.Count > 0 Then
                        Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                        Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt2.MessageForAdding = vbLf & "Select the block to be populated:"
                        Object_Prompt2.SingleOnly = True
                        Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                        If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            If IsNothing(Rezultat2) = False Then
                                Dim Ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForWrite)
                                If TypeOf Ent2 Is BlockReference Then
                                    Dim Block2 As BlockReference = Ent2
                                    If Block2.AttributeCollection.Count > 0 Then
                                        Dim Data_table_Tags As New System.Data.DataTable
                                        Data_table_Tags.Columns.Add("TAG", GetType(String))
                                        Data_table_Tags.Columns.Add("X", GetType(Double))
                                        Data_table_Tags.Columns.Add("Y", GetType(Double))
                                        Dim Index3 As Integer = 0

                                        For Each atrID In Block2.AttributeCollection
                                            Dim Atr2 As AttributeReference = atrID.GetObject(OpenMode.ForWrite)
                                            Data_table_Tags.Rows.Add()
                                            Data_table_Tags.Rows(Index3).Item("TAG") = Atr2.Tag
                                            If Atr2.IsMTextAttribute = True Then
                                                Atr2.MTextAttribute.Contents = ""
                                            Else
                                                Atr2.TextString = ""
                                            End If
                                            Data_table_Tags.Rows(Index3).Item("X") = Round(Atr2.Position.X, 0)
                                            Data_table_Tags.Rows(Index3).Item("Y") = Round(Atr2.Position.Y, 0)
                                            Index3 = Index3 + 1
                                        Next

                                        '" DESC,"  " ASC"
                                        Data_table_Tags = Sort_data_table_2_columns(Data_table_Tags, "Y", " DESC,", "X", " ASC")
                                        Dim Colectie_valori As New Specialized.StringCollection
                                        Dim Colectie_shrink As New IntegerCollection

                                        For k = 0 To Colectie_dwg.Rows.Count - 1
                                            Dim Descr2 As String = Colectie_dwg.Rows(k).Item("DESCR")

                                            If Strings.Len(Descr2) <= Nr_max_litere_descr Then
                                                Colectie_valori.Add(Colectie_dwg.Rows(k).Item("DWG"))
                                                If Len(Colectie_dwg.Rows(k).Item("DWG")) > Nr_max_litere_dwg Then
                                                    Colectie_shrink.Add(1)
                                                Else
                                                    Colectie_shrink.Add(0)
                                                End If
                                                Colectie_valori.Add(Colectie_dwg.Rows(k).Item("DESCR"))
                                                Colectie_shrink.Add(0)
                                            End If


                                            If Strings.Len(Descr2) > Nr_max_litere_descr Then
                                                Dim DescrA As String = Strings.Left(Descr2, Nr_max_litere_descr)
                                                Dim Split_car As Integer
                                                For l = 1 To Nr_max_litere_descr
                                                    If Strings.Left(Strings.Right(DescrA, l), 1) = " " Then
                                                        Split_car = Strings.Len(DescrA) - l
                                                        Exit For
                                                    End If
                                                Next
                                                Colectie_valori.Add(Colectie_dwg.Rows(k).Item("DWG"))
                                                If Len(Colectie_dwg.Rows(k).Item("DWG")) > Nr_max_litere_dwg Then
                                                    Colectie_shrink.Add(1)
                                                Else
                                                    Colectie_shrink.Add(0)
                                                End If
                                                Colectie_valori.Add(Strings.Left(Descr2, Split_car))
                                                Colectie_shrink.Add(0)

                                                Dim DescrB As String = Strings.Right(Descr2, Strings.Len(Descr2) - Split_car - 1)

                                                If Strings.Len(DescrB) <= Nr_max_litere_descr Then
                                                    Colectie_valori.Add(" ")
                                                    Colectie_shrink.Add(0)
                                                    Colectie_valori.Add(DescrB)
                                                    Colectie_shrink.Add(0)
                                                Else
                                                    Colectie_valori.Add(" ")
                                                    Colectie_shrink.Add(0)
                                                    Dim DescrC As String = Strings.Left(DescrB, Nr_max_litere_descr)
                                                    Dim Split_carB As Integer
                                                    For l = 1 To Nr_max_litere_descr
                                                        If Strings.Left(Strings.Right(DescrC, l), 1) = " " Then
                                                            Split_carB = Strings.Len(DescrC) - l
                                                            Exit For
                                                        End If
                                                    Next
                                                    Colectie_valori.Add(Strings.Left(DescrB, Split_carB))
                                                    Colectie_shrink.Add(0)


                                                    Dim DescrB1 As String = Strings.Right(DescrB, Strings.Len(DescrB) - Split_carB - 1)

                                                    If Strings.Len(DescrB1) <= Nr_max_litere_descr Then
                                                        Colectie_valori.Add(" ")
                                                        Colectie_shrink.Add(0)
                                                        Colectie_valori.Add(DescrB1)
                                                        Colectie_shrink.Add(0)
                                                    Else
                                                        Colectie_valori.Add(" ")
                                                        Colectie_shrink.Add(0)
                                                        Dim DescrC1 As String = Strings.Left(DescrB1, Nr_max_litere_descr)
                                                        Dim Split_carB1 As Integer
                                                        For l = 1 To Nr_max_litere_descr
                                                            If Strings.Left(Strings.Right(DescrC1, l), 1) = " " Then
                                                                Split_carB1 = Strings.Len(DescrC1) - l
                                                                Exit For
                                                            End If
                                                        Next
                                                        Colectie_valori.Add(Strings.Left(DescrB1, Split_carB1))
                                                        Colectie_shrink.Add(0)
                                                        Dim DescrB2 As String = Strings.Right(DescrB1, Strings.Len(DescrB1) - Split_carB1 - 1)

                                                        Colectie_valori.Add(" ")
                                                        Colectie_shrink.Add(0)
                                                        Colectie_valori.Add(DescrB2)
                                                        Colectie_shrink.Add(0)
                                                    End If
                                                End If
                                            End If
                                        Next

                                        Dim Index_colectie As Integer = 0


                                        For Each atrID In Block2.AttributeCollection

                                            Dim Atr3 As AttributeReference = atrID.GetObject(OpenMode.ForWrite)
                                            Dim Tag3 As String = Atr3.Tag

                                            For m = 0 To Data_table_Tags.Rows.Count - 1
                                                If Tag3.ToUpper = Data_table_Tags.Rows(m).Item("TAG").ToString.ToUpper Then
                                                    Index_colectie = m
                                                    Exit For
                                                End If

                                            Next

                                            If Index_colectie <= Colectie_shrink.Count - 1 Then
                                                If Colectie_shrink(Index_colectie) = 1 Then
                                                    Atr3.WidthFactor = 0.9
                                                Else
                                                    Atr3.WidthFactor = 1
                                                End If

                                                If Atr3.IsMTextAttribute = True Then
                                                    Atr3.MTextAttribute.Contents = Colectie_valori(Index_colectie)
                                                Else
                                                    Atr3.TextString = Colectie_valori(Index_colectie)
                                                End If
                                            End If



                                        Next

                                        If Colectie_valori.Count > Block2.AttributeCollection.Count Then
                                            MsgBox("You have more values than block attributes for this sheet!")
                                        End If


                                    End If
                                End If
                            End If
                        End If

                    End If


                    Trans1.Commit()
                End Using




                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub




    Private Sub Button_LOAD_TEXT_FROM_EXCEL_USA_Click(sender As Object, e As EventArgs) Handles Button_LOAD_TEXT_FROM_EXCEL_USA.Click
        Try



            If TextBox_DWG_NAME.Text = "" Or TextBox_DWG_NO.Text = "" Then
                MsgBox("Please specify the EXCEL COLUMN!")
                Exit Sub
            End If
            If TextBox_row_per_block.Text = "" Then
                MsgBox("Please specify the number of lines per block")
                Exit Sub
            End If
            If IsNumeric(TextBox_row_per_block.Text) = False Then
                MsgBox("Please specify the number of lines per block")
                Exit Sub
            End If

            If TextBox_excel_spacing.Text = "" Then
                MsgBox("Please specify the excel spacing")
                Exit Sub
            End If
            If IsNumeric(TextBox_excel_spacing.Text) = False Then
                MsgBox("Please specify the excel spacing")
                Exit Sub
            End If


            If TextBox_start_USA.Text = "" Then
                MsgBox("Please specify the EXCEL START ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_start_USA.Text) = False Then
                With TextBox_start_USA
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If
            If TextBox_end_USA.Text = "" Then
                MsgBox("Please specify the EXCEL END ROW!")
                Exit Sub
            End If

            If IsNumeric(TextBox_end_USA.Text) = False Then
                With TextBox_end_USA
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify END row")

                Exit Sub
            End If

            If Val(TextBox_start_USA.Text) < 1 Then
                With TextBox_start_USA
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Start row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_end_USA.Text) < 1 Then
                With TextBox_end_USA
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row can't be smaller than 1")

                Exit Sub
            End If

            If Val(TextBox_end_USA.Text) < Val(TextBox_start_USA.Text) Then
                MsgBox("END row smaller than start row")

                Exit Sub
            End If

            Dim NR_rand_block As Integer = Abs(CInt(TextBox_row_per_block.Text))
            Dim Spacing_excel As Integer = Abs(CInt(TextBox_excel_spacing.Text))
            Dim start1 As Integer = CInt(TextBox_start_USA.Text)
            Dim end1 As Integer = CInt(TextBox_end_USA.Text)
            If Not ((end1 - start1 + 1) + Spacing_excel) / (NR_rand_block + Spacing_excel) = CInt(((end1 - start1 + 1) + Spacing_excel) / (NR_rand_block + Spacing_excel)) Then
                MsgBox("You have to have a selection multiple of no of rows and rows spacing in excel")

                Exit Sub
            End If

            Nr_blocks = ((end1 - start1 + 1) + Spacing_excel) / (NR_rand_block + Spacing_excel)


            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

            Data_table_REF_USA = New System.Data.DataTable
            Data_table_REF_USA.Columns.Add("INDEX", GetType(Integer))
            Data_table_REF_USA.Columns.Add("DWG_NO", GetType(String))
            Data_table_REF_USA.Columns.Add("DWG_NAME", GetType(String))
            Data_table_REF_USA.Columns.Add("ALIGNMENT", GetType(String))

            Dim Index As Integer = 0

            Dim Nr1 As Integer = 0


            For i = start1 To end1


                Nr1 = Nr1 + 1

                If Nr1 = NR_rand_block + 1 Then Nr1 = -1
                'MsgBox(Nr1)

                If Nr1 <= NR_rand_block And Nr1 > 0 Then
                    Dim DWG_NO As String = " "
                    Dim DWG_NAME As String = " "
                    Dim ALIGNMENT As String = " "

                    If Not Replace(W1.Range(TextBox_DWG_NO.Text.ToUpper & i).Text, " ", "") = "" Then
                        DWG_NO = W1.Range(TextBox_DWG_NO.Text.ToUpper & i).Value2
                    End If

                    If Not Replace(W1.Range(TextBox_DWG_NAME.Text.ToUpper & i).Text, " ", "") = "" Then
                        DWG_NAME = W1.Range(TextBox_DWG_NAME.Text.ToUpper & i).Value2
                    End If

                    If i - 1 > 0 Then
                        If Not Replace(W1.Range(TextBox_DWG_NO.Text.ToUpper & i - Nr1).Text, " ", "") = "" Then
                            ALIGNMENT = W1.Range(TextBox_DWG_NO.Text.ToUpper & i - Nr1).Value2
                        End If
                    End If



                    Data_table_REF_USA.Rows.Add()
                    Dim Index_dwg As Integer


                    For k = 1 To Nr_blocks
                        If i <= start1 + k * NR_rand_block + (k - 1) * Spacing_excel Then
                            Index_dwg = k
                            Exit For
                        End If
                    Next





                    Data_table_REF_USA.Rows(Index).Item("INDEX") = Abs(Index_dwg)
                    Data_table_REF_USA.Rows(Index).Item("DWG_NO") = DWG_NO
                    Data_table_REF_USA.Rows(Index).Item("DWG_NAME") = DWG_NAME
                    If Not ALIGNMENT = " " Then Data_table_REF_USA.Rows(Index).Item("ALIGNMENT") = ALIGNMENT


                    Index = Index + 1

                End If



            Next

            Dim String1 As String = ""
            For i = 0 To Data_table_REF_USA.Rows.Count - 1
                String1 = String1 & vbCrLf & Data_table_REF_USA.Rows(i).Item("INDEX") & Chr(9) & Data_table_REF_USA.Rows(i).Item("DWG_NO") & Chr(9) & Data_table_REF_USA.Rows(i).Item("DWG_NAME")
            Next
            My.Computer.Clipboard.SetText(String1)

            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_add_to_blocks_USA_Click(sender As Object, e As EventArgs) Handles Button_add_to_blocks_USA.Click
        Try
            If Data_table_REF_USA.Rows.Count = 0 Then
                MsgBox("You don't have loaded any references!")
                Exit Sub
            End If



            If IsNumeric(TextBox_blocks_spacing.Text) = False Then
                With TextBox_blocks_spacing
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify the distance between blocks!")

                Exit Sub
            End If




            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Creaza_layer("NO PLOT", 40, "NO PLOT", False)
                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select the block to be populated:"
                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                    If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat2) = False Then
                            Dim Ent2 As Entity = Rezultat2.Value.Item(0).ObjectId.GetObject(OpenMode.ForWrite)
                            If TypeOf Ent2 Is BlockReference Then
                                Dim Block2 As BlockReference = Ent2
                                If Block2.AttributeCollection.Count > 0 Then
                                    Dim Data_table_Tags As New System.Data.DataTable
                                    Data_table_Tags.Columns.Add("TAG", GetType(String))
                                    Data_table_Tags.Columns.Add("X", GetType(Double))
                                    Data_table_Tags.Columns.Add("Y", GetType(Double))
                                    Dim Index3 As Integer = 0
                                    For Each atrID In Block2.AttributeCollection
                                        Dim Atr2 As AttributeReference = atrID.GetObject(OpenMode.ForWrite)
                                        Data_table_Tags.Rows.Add()
                                        Data_table_Tags.Rows(Index3).Item("TAG") = Atr2.Tag
                                        Data_table_Tags.Rows(Index3).Item("X") = Round(Atr2.Position.X, 0)
                                        Data_table_Tags.Rows(Index3).Item("Y") = Round(Atr2.Position.Y, 0)
                                        Index3 = Index3 + 1
                                        If Atr2.IsMTextAttribute = True Then
                                            Atr2.MTextAttribute.Contents = ""
                                        Else
                                            Atr2.TextString = ""
                                        End If
                                    Next

                                    '" DESC,"  " ASC"

                                    Data_table_Tags = Sort_data_table_2_columns(Data_table_Tags, "Y", " DESC,", "X", " ASC")
                                    Dim Colectie_valori As New Specialized.StringCollection

                                    For k = 0 To Data_table_REF_USA.Rows.Count - 1
                                        If Data_table_REF_USA.Rows(k).Item("INDEX") = 1 Then
                                            Colectie_valori.Add(Data_table_REF_USA.Rows(k).Item("DWG_NO"))
                                            Colectie_valori.Add(Data_table_REF_USA.Rows(k).Item("DWG_NAME"))
                                        End If
                                    Next

                                    Dim Index_colectie As Integer = 0

                                    For Each atrID In Block2.AttributeCollection

                                        Dim Atr3 As AttributeReference = atrID.GetObject(OpenMode.ForWrite)
                                        Dim Tag3 As String = Atr3.Tag



                                        For m = 0 To Data_table_Tags.Rows.Count - 1
                                            If Tag3.ToUpper = Data_table_Tags.Rows(m).Item("TAG").ToString.ToUpper Then
                                                Index_colectie = m
                                                Exit For
                                            End If

                                        Next


                                        If Atr3.IsMTextAttribute = True Then
                                            Atr3.MTextAttribute.Contents = Colectie_valori(Index_colectie)
                                        Else
                                            Atr3.TextString = Colectie_valori(Index_colectie)
                                        End If

                                    Next
                                    If Colectie_valori.Count > Block2.AttributeCollection.Count Then
                                        MsgBox("You have more values than block attributes for this sheet!")
                                    End If


                                    If IsDBNull(Data_table_REF_USA.Rows(0).Item("ALIGNMENT")) = False Then


                                        Dim Text1 As New DBText
                                        Text1.TextString = Data_table_REF_USA.Rows(0).Item("ALIGNMENT")

                                        Text1.Justify = AttachmentPoint.MiddleRight
                                        Text1.AlignmentPoint = New Point3d(Block2.Position.X - 50, Block2.Position.Y - 0.5 * CDbl(TextBox_blocks_spacing.Text), 0)

                                        Text1.Height = 50
                                        Text1.Layer = "NO PLOT"
                                        BTrecord.AppendEntity(Text1)
                                        Trans1.AddNewlyCreatedDBObject(Text1, True)

                                    End If


                                    If Nr_blocks > 1 Then
                                        For i = 2 To Nr_blocks
                                            Colectie_valori = New Specialized.StringCollection
                                            Dim Alignment As String = ""

                                            Dim exists1 As Boolean = False

                                            For k = 0 To Data_table_REF_USA.Rows.Count - 1
                                                If Data_table_REF_USA.Rows(k).Item("INDEX") = i Then
                                                    Colectie_valori.Add(Data_table_REF_USA.Rows(k).Item("DWG_NO"))
                                                    Colectie_valori.Add(Data_table_REF_USA.Rows(k).Item("DWG_NAME"))
                                                    If exists1 = False Then
                                                        If IsDBNull(Data_table_REF_USA.Rows(k).Item("ALIGNMENT")) = False Then
                                                            Alignment = Data_table_REF_USA.Rows(k).Item("ALIGNMENT")
                                                            exists1 = True
                                                        End If
                                                    End If
                                                End If
                                            Next

                                            Dim Brec As BlockReference
                                            Brec = InsertBlock_with_multiple_atributes("", Block2.Name, New Point3d(Block2.Position.X, Block2.Position.Y - (i - 1) * CDbl(TextBox_blocks_spacing.Text), 0), 1, BTrecord, Block2.Layer, Colectie_valori, Colectie_valori)

                                            If Not Alignment = "" Then


                                                Dim Text1 As New DBText
                                                Text1.TextString = Alignment

                                                Text1.Justify = AttachmentPoint.MiddleRight
                                                Text1.AlignmentPoint = New Point3d(Block2.Position.X - 50, Block2.Position.Y - (i - 1) * CDbl(TextBox_blocks_spacing.Text) - 0.5 * CDbl(TextBox_blocks_spacing.Text), 0)

                                                Text1.Height = 50
                                                Text1.Layer = "NO PLOT"
                                                BTrecord.AppendEntity(Text1)
                                                Trans1.AddNewlyCreatedDBObject(Text1, True)

                                            End If



                                            For Each atrID In Brec.AttributeCollection

                                                Dim Atr3 As AttributeReference = atrID.GetObject(OpenMode.ForWrite)
                                                Dim Tag3 As String = Atr3.Tag



                                                For m = 0 To Data_table_Tags.Rows.Count - 1
                                                    If Tag3.ToUpper = Data_table_Tags.Rows(m).Item("TAG").ToString.ToUpper Then
                                                        Index_colectie = m
                                                        Exit For
                                                    End If

                                                Next


                                                If Atr3.IsMTextAttribute = True Then
                                                    Atr3.MTextAttribute.Contents = Colectie_valori(Index_colectie)
                                                Else
                                                    Atr3.TextString = Colectie_valori(Index_colectie)
                                                End If

                                            Next


                                        Next



                                    End If


                                End If
                            End If
                        End If
                    End If






                    Trans1.Commit()
                End Using


                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using

        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_write_to_excel_Click(sender As Object, e As EventArgs) Handles Button_write_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Try
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select blocks:"

                    Object_Prompt.SingleOnly = False

                    Rezultat1 = Editor1.GetSelection(Object_Prompt)


                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        Freeze_operations = False
                        Exit Sub
                    End If



                    Dim Row1 As Integer = 1
                    If IsNumeric(TextBox_transfer_start_row.Text) = True Then
                        Row1 = Abs(CInt(TextBox_transfer_start_row.Text))
                    End If
                    Dim Column_atr_name As String
                    Column_atr_name = TextBox_col_atr_name.Text
                    Dim Column_atr_value As String
                    Column_atr_value = TextBox_col_atr_value_start.Text

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        If IsNothing(Rezultat1) = False Then
                            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                                    For i = 0 To Rezultat1.Value.Count - 1



                                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                        Obj1 = Rezultat1.Value.Item(i)
                                        Dim Ent1 As Entity
                                        Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForRead)

                                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                            Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                            If Block1.AttributeCollection.Count > 0 Then
                                                W1.Range(Column_atr_name & Row1).Value2 = Block1.Name
                                                Row1 = Row1 + 1
                                                For Each id As ObjectId In Block1.AttributeCollection
                                                    If Not id.IsErased Then
                                                        Dim attRef As AttributeReference = DirectCast(Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                                        Dim Continut As String = attRef.TextString
                                                        Dim Tag As String = attRef.Tag
                                                        W1.Range(Column_atr_value & Row1).Value2 = Continut
                                                        W1.Range(Column_atr_name & Row1).Value2 = Tag
                                                        Row1 = Row1 + 1
                                                    End If
                                                Next


                                            End If


                                        End If



                                    Next
                                    Trans1.Commit()

                                End Using
                            End Using
                        End If
                    End If



                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                Catch ex As Exception
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    MsgBox(ex.Message)
                End Try
            Catch EX As System.Runtime.InteropServices.COMException
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(EX.Message)
            End Try

            Freeze_operations = False
        End If

    End Sub

    Private Sub Button_add_to_blocks_Click(sender As Object, e As EventArgs) Handles Button_add_to_blocks.Click
        If Freeze_operations = False Then
            Freeze_operations = True


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select block:"

                Object_Prompt.SingleOnly = True

                Rezultat1 = Editor1.GetSelection(Object_Prompt)


                If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub
                End If



                Dim Start1 As Integer = 1
                If IsNumeric(TextBox_transfer_row_start.Text) = True Then
                    Start1 = Abs(CInt(TextBox_transfer_row_start.Text))
                End If
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_transfer_row_end.Text) = True Then
                    End1 = Abs(CInt(TextBox_transfer_row_end.Text))
                End If

                If End1 < Start1 Then

                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub

                End If




                Dim Column_start As String = TextBox_block_att_value_column_start.Text.ToUpper
                Dim Column_end As String = TextBox_block_att_value_column_end.Text.ToUpper
                Dim Column_atr_name As String = TextBox_block_att_name_column.Text.ToUpper


                Dim Col_start As Integer = 0
                Dim Col_end As Integer = 0
                Dim Col_atr As Integer = 0

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Nothing


                Try
                    W1 = Get_active_worksheet_from_Excel_with_error()
                    If IsNothing(W1) = False Then
                        Col_start = W1.Range(Column_start & "1").Column
                        Col_end = W1.Range(Column_end & "1").Column
                        Col_atr = W1.Range(Column_atr_name & "1").Column
                    End If

                Catch ex As System.SystemException

                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub

                End Try

                Dim Row_with_layout_names As Integer = 0
                If IsNumeric(TextBox_ROW_layout_NAME.Text) = True Then
                    Row_with_layout_names = CInt(TextBox_ROW_layout_NAME.Text)
                End If


                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument
                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                Dim Layoutdict As DBDictionary = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)
                                Dim Ent1 As Entity
                                Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)
                                Dim Block_name As String = ""



                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                    Dim Block1 As BlockReference = TryCast(Ent1, BlockReference)
                                    If IsNothing(Block1) = False Then

                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim BlockTrec As BlockTableRecord = Nothing
                                            If Block1.IsDynamicBlock = True Then
                                                BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                Block_name = BlockTrec.Name
                                            Else
                                                BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                Block_name = BlockTrec.Name
                                            End If
                                        End If
                                    End If

                                End If

                                Dim Data_table1 As New System.Data.DataTable
                                Data_table1.Columns.Add("ATR", GetType(String))
                                Data_table1.Columns.Add("VALUE", GetType(String))
                                Data_table1.Columns.Add("LAYOUT", GetType(String))
                                Dim iNDEX1 As Integer = 0

                                For j = Col_start To Col_end
                                    Dim Layout_excel As String = W1.Cells(Row_with_layout_names, j).Value2





                                    For i = Start1 To End1
                                        Dim Attrib_name_excel As String = W1.Range(Column_atr_name & i.ToString).Value2
                                        Dim Value_excel As String = W1.Cells(i, j).Value2
                                        If Not Attrib_name_excel = "" And Not Layout_excel = "" Then
                                            Data_table1.Rows.Add()
                                            Data_table1.Rows(iNDEX1).Item("ATR") = Attrib_name_excel
                                            Data_table1.Rows(iNDEX1).Item("VALUE") = Value_excel
                                            Data_table1.Rows(iNDEX1).Item("LAYOUT") = Layout_excel
                                            iNDEX1 = iNDEX1 + 1
                                        End If


                                    Next
                                Next

                                If Data_table1.Rows.Count > 0 Then
                                    Dim Lay_name As String = ""
                                    For i = 0 To Data_table1.Rows.Count - 1
                                        If IsDBNull(Data_table1.Rows(i).Item("LAYOUT")) = False Then
                                            If Not Lay_name = Data_table1.Rows(i).Item("LAYOUT") Then
                                                Lay_name = Data_table1.Rows(i).Item("LAYOUT")

                                                For Each entry As DBDictionaryEntry In Layoutdict
                                                    Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)



                                                    If Not Layout1.TabOrder = 0 And Layout1.LayoutName = Lay_name Then
                                                        LayoutManager1.CurrentLayout = Lay_name

                                                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                                                        For Each obid As ObjectId In BTrecord
                                                            Dim Block1 As BlockReference = TryCast(Trans1.GetObject(obid, OpenMode.ForRead), BlockReference)
                                                            If IsNothing(Block1) = False Then
                                                                If Block1.Name = Block_name Then
                                                                    If Block1.AttributeCollection.Count > 0 Then

                                                                        For Each id As ObjectId In Block1.AttributeCollection
                                                                            If Not id.IsErased Then
                                                                                Dim attRef As AttributeReference = DirectCast(Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), AttributeReference)


                                                                                Dim Tag As String = attRef.Tag
                                                                                For s = 0 To Data_table1.Rows.Count - 1
                                                                                    If IsDBNull(Data_table1.Rows(s).Item("ATR")) = False And IsDBNull(Data_table1.Rows(s).Item("LAYOUT")) = False Then
                                                                                        If Data_table1.Rows(s).Item("LAYOUT") = Lay_name Then
                                                                                            Dim Attrib_name As String = Data_table1.Rows(s).Item("ATR")
                                                                                            Dim Value1 As String = ""

                                                                                            If IsDBNull(Data_table1.Rows(s).Item("VALUE")) = False Then
                                                                                                Value1 = Data_table1.Rows(s).Item("VALUE")
                                                                                            End If

                                                                                            If Tag.ToUpper = Attrib_name.ToUpper Then
                                                                                                attRef.TextString = Value1
                                                                                                If attRef.IsMTextAttribute = False Then

                                                                                                Else
                                                                                                    'attRef.MTextAttribute.Contents = Value1
                                                                                                End If

                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                            End If
                                                                        Next
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                End If













                                Trans1.Commit()

                            End Using
                        End Using
                    End If
                End If



                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try


            Freeze_operations = False
        End If



    End Sub

    Public Function Get_active_worksheet_from_Excel_with_error() As Microsoft.Office.Interop.Excel.Worksheet
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook

        Try
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch ex As System.Exception
            Return Nothing
        Finally
            'Excel1.ActiveWindow.DisplayGridlines = True
            If Excel1.Workbooks.Count = 0 Then Excel1.Workbooks.Add()
            If Excel1.Visible = False Then Excel1.Visible = True
            Workbook1 = Excel1.ActiveWorkbook
        End Try
        Return Workbook1.ActiveSheet
    End Function



    Private Sub Button_clear_lists_Click(sender As Object, e As EventArgs) Handles Button_clear_lists.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count > 0 Then
                ListBox_DWG.Items.Clear()
            End If
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_remove_items_list_Click(sender As Object, e As EventArgs) Handles Button_remove_items_list.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count > 0 Then
                If ListBox_DWG.SelectedIndex >= 0 Then
                    ListBox_DWG.Items.RemoveAt((ListBox_DWG.SelectedIndex))
                End If
            End If
            Freeze_operations = False
        End If
    End Sub
    Private Sub Button_load_DWG_Click(sender As Object, e As EventArgs) Handles Button_load_DWG.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Drawing Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
                FileBrowserDialog1.Multiselect = True

                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    For Each file1 In FileBrowserDialog1.FileNames
                        ListBox_DWG.Items.Add(file1)
                    Next

                End If

            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_redefine_block_Click(sender As Object, e As EventArgs) Handles Button_redefine_block.Click



        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Try

                Try

                    If IsNumeric(TextBox_Layout_index_START.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If
                    If IsNumeric(TextBox_Layout_index_END.Text) = False Then
                        MsgBox("Please specify the layout for block insertion!")
                        Freeze_operations = False
                        Exit Sub
                    End If



                    Dim Column_start As String = TextBox_ATRIB_START.Text.ToUpper
                    Dim Column_end As String = TextBox_ATRIB_END.Text.ToUpper
                    Dim Column_atr_name As String = TextBox_ATRIB_NAME.Text.ToUpper

                    Dim Row_with_file_names As Integer = 0
                    If IsNumeric(TextBox_ROW_FILE_NAME.Text) = True Then
                        Row_with_file_names = CInt(TextBox_ROW_FILE_NAME.Text)
                    End If


                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                    W1 = Nothing
                    Dim Col_start As Integer = 0
                    Dim Col_end As Integer = 0
                    Dim Col_atr As Integer = 0


                    Try
                        W1 = Get_active_worksheet_from_Excel_with_error()
                        If IsNothing(W1) = False Then
                            Col_start = W1.Range(Column_start & "1").Column
                            Col_end = W1.Range(Column_end & "1").Column
                            Col_atr = W1.Range(Column_atr_name & "1").Column
                        End If

                    Catch ex As System.SystemException
                        W1 = Nothing
                        Col_start = 0
                        Col_end = 0
                        Col_atr = 0
                    End Try


                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()




                                Dim Name_of_the_block As String = TextBox_block_name.Text



                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then

                                        Dim Colectie_nume_atribute As New Specialized.StringCollection
                                        Dim Colectie_valori_atribute As New Specialized.StringCollection

                                        If IsNothing(W1) = False Then
                                            Dim Start1 As Integer = 0
                                            Dim End1 As Integer = 0
                                            If IsNumeric(TextBox_blocks_att_row_start.Text) = True Then
                                                Start1 = CInt(TextBox_blocks_att_row_start.Text)
                                            End If
                                            If IsNumeric(TextBox_blocks_att_row_end.Text) = True Then
                                                End1 = CInt(TextBox_blocks_att_row_end.Text)
                                            End If

                                            If Not Row_with_file_names = 0 And Not Col_start = 0 And Not Col_end = 0 And Not Col_end < Col_start And Not Col_atr = 0 And Not Start1 = 0 And Not End1 = 0 And Not End1 < Start1 Then
                                                For j = Col_start To Col_end
                                                    Dim Excel_file As String = W1.Cells(Row_with_file_names, j).Value2
                                                    If Drawing1.ToUpper.Contains(Excel_file.ToUpper) = True Then

                                                        For k = Start1 To End1

                                                            Dim Atr_name As String = ""
                                                            Atr_name = W1.Cells(k, Col_atr).Value2
                                                            Dim Atr_value As String = ""
                                                            Atr_value = W1.Cells(k, j).Value2

                                                            If Not Atr_name = "" Then
                                                                Colectie_nume_atribute.Add(Atr_name)
                                                                Colectie_valori_atribute.Add(Atr_value)
                                                            End If

                                                        Next

                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If


                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1





                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                            Dim Index_datatable As Integer = 0


                                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                                            If BlockTable1.Has(Name_of_the_block) = True Then
                                                For Each entry As DBDictionaryEntry In Layoutdict
                                                    Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                    If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) Then
                                                        Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                        UpdateBlock_with_multiple_atributes(Name_of_the_block, Database1, BTrecord, Colectie_nume_atribute, Colectie_valori_atribute)
                                                    End If
                                                Next
                                            End If

                                            Trans1.Commit()

                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf & _
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_visretain_Click(sender As Object, e As EventArgs) Handles Button_visretain.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            If RadioButton_visretain_0.Checked = False And RadioButton_visretain_1.Checked = False Then
                MsgBox("please choose a visretain value")
                Freeze_operations = False
                Exit Sub
            End If

            If ListBox_DWG.Items.Count = 0 Then
                MsgBox("please specify the drawings!")
                Freeze_operations = False
                Exit Sub
            End If

            Dim Error_list As String = ""

            Try

                Try




                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()




                                Dim Name_of_the_block As String = TextBox_block_name.Text



                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then



                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1





                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                            If RadioButton_visretain_0.Checked = True Then
                                                Database1.Visretain = False
                                            Else
                                                Database1.Visretain = True
                                            End If

                                            Trans1.Commit()

                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using

                                        Database1.Dispose()

                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If

123:
                                Next

                                Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf & _
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_delete_AecObjExplode_Click(sender As Object, e As EventArgs) Handles Button_delete_AecObjExplode.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count = 0 Then
                MsgBox("please specify the drawings!")
                Freeze_operations = False
                Exit Sub
            End If
            Dim Error_list As String = ""

            Try
                Try
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument
                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                                Dim Name_of_the_block As String = TextBox_block_name.Text
                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then
                                        Dim Database1 As New Database(False, True)
                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try
                                        HostApplicationServices.WorkingDatabase = Database1
                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                                            Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                                            LayerTable1 = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                            Dim Layer0_id As Autodesk.AutoCAD.DatabaseServices.ObjectId = LayerTable1.Item("0")
                                            Database1.Clayer = Layer0_id
                                            Dim Lista_del As New Specialized.StringCollection
                                            For Each Lid As ObjectId In LayerTable1
                                                Dim LayerTrec As LayerTableRecord = Trans1.GetObject(Lid, OpenMode.ForRead)
                                                If LayerTrec.Name.ToLower.Contains("aecobjexplode_") = True Then
                                                    Lista_del.Add(LayerTrec.Name)
                                                End If
                                            Next
                                            If Lista_del.Count > 0 Then
                                                Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                                Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                                For Each entry As DBDictionaryEntry In Layoutdict
                                                    Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                    Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                    For Each Id1 As ObjectId In BTrecord
                                                        Dim ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                        If IsNothing(ent1) = False Then
                                                            If Lista_del.Contains(ent1.Layer) = True Then
                                                                ent1.UpgradeOpen()
                                                                ent1.Erase()
                                                            End If
                                                        End If
                                                    Next
                                                Next
                                                Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                                For Each id1 As ObjectId In BlockTable1
                                                    Dim bL1 As BlockTableRecord = Trans1.GetObject(id1, OpenMode.ForRead)
                                                    If Lista_del.Contains(bL1.Name) = True Then
                                                        bL1.UpgradeOpen()
                                                        bL1.Erase()
                                                    End If
                                                Next
                                                For Each layername As String In Lista_del
                                                    Dim Layer1 As LayerTableRecord = Trans1.GetObject(LayerTable1(layername), OpenMode.ForWrite)
                                                    Layer1.Erase()
                                                Next
                                            End If
                                            Trans1.Commit()
                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using
                                        Database1.Dispose()
                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If
123:
                                Next
                                Trans11.Commit()
                            End Using
                        End Using
                    End If
                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf & _
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_split_layouts_Click(sender As Object, e As EventArgs) Handles Button_split_layouts.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Error_list As String = ""

            Try

                Try







                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                                Dim Colectie_nume_fisiere As New Specialized.StringCollection



                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then


                                        Dim Database1 As New Database(False, True)

                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try


                                        HostApplicationServices.WorkingDatabase = Database1





                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction


                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary

                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)

                                            Dim Path1 As String = IO.Path.GetDirectoryName(Drawing1)
                                            If Not Strings.Right(Path1, 1) = "\" Then
                                                Path1 = Path1 & "\"
                                            End If

                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                If Layout1.TabOrder > 0 Then
                                                    Colectie_nume_fisiere.Add(Path1 & Layout1.LayoutName & ".dwg")
                                                End If
                                            Next



                                            Trans1.Commit()

                                            For j = 0 To Colectie_nume_fisiere.Count - 1
                                                Dim File1 As String = IO.Path.GetFileNameWithoutExtension(Colectie_nume_fisiere(j))

                                                Dim index1 As Integer = 1
                                                Do Until IO.File.Exists(Path1 & File1 & ".dwg") = False
                                                    File1 = File1 & index1
                                                    Colectie_nume_fisiere(j) = Path1 & File1 & ".dwg"
                                                Loop
                                                Try
                                                    Database1.SaveAs(Colectie_nume_fisiere(j), True, DwgVersion.Current, Database1.SecurityParameters)
                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Colectie_nume_fisiere(j)
                                                End Try
                                            Next

                                        End Using



                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database








                                    End If




123:
                                Next


                                For j = 0 To Colectie_nume_fisiere.Count - 1

                                    Dim Database2 As New Database(False, True)

                                    Try
                                        Try
                                            Try
                                                Database2.ReadDwgFile(Colectie_nume_fisiere(j), FileOpenMode.OpenForReadAndAllShare, False, "")
                                            Catch ex As Exception
                                                MsgBox(Colectie_nume_fisiere(j) & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Colectie_nume_fisiere(j)
                                                GoTo 124
                                            End Try
                                        Catch ex As IO.IOException
                                            MsgBox(Colectie_nume_fisiere(j) & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Colectie_nume_fisiere(j)
                                            GoTo 124
                                        End Try
                                    Catch ex As System.SystemException
                                        MsgBox(Colectie_nume_fisiere(j) & vbCrLf & "could not be open")
                                        Error_list = Error_list & vbCrLf & Colectie_nume_fisiere(j)
                                        GoTo 124
                                    End Try


                                    HostApplicationServices.WorkingDatabase = Database2





                                    Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database2.TransactionManager.StartTransaction

                                        Dim LayoutManager2 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                        Dim Layoutdict As DBDictionary

                                        Layoutdict = Trans2.GetObject(Database2.LayoutDictionaryId, OpenMode.ForRead)

                                        Dim File2 As String = IO.Path.GetFileNameWithoutExtension(Colectie_nume_fisiere(j))
                                        Dim Colectie_del As New Specialized.StringCollection

                                        Dim Fisier1 As String = IO.Path.GetFileNameWithoutExtension(Colectie_nume_fisiere(j))

                                        For Each entry As DBDictionaryEntry In Layoutdict
                                            Dim Layout2 As Layout = Trans2.GetObject(LayoutManager2.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                            If Not Layout2.LayoutName = Fisier1 And Layout2.TabOrder > 0 Then
                                                Colectie_del.Add(Layout2.LayoutName)
                                            End If
                                        Next

                                        For k = 0 To Colectie_del.Count - 1
                                            LayoutManager2.DeleteLayout(Colectie_del(k))
                                        Next


                                        Trans2.Commit()

                                        Try
                                            Database2.SaveAs(Colectie_nume_fisiere(j), True, DwgVersion.Current, Database2.SecurityParameters)
                                        Catch ex As Exception

                                        End Try

                                    End Using



                                    HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                Next

124:
                                ' Trans11.Commit()
                            End Using
                        End Using
                    End If


                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf & _
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try

            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False

        End If
    End Sub

    Private Sub Button_fix_PL_Click(sender As Object, e As EventArgs) Handles Button_fix_PL.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            If ListBox_DWG.Items.Count = 0 Then
                MsgBox("please specify the drawings!")
                Freeze_operations = False
                Exit Sub
            End If
            Dim Error_list As String = ""

            Try
                Try
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument
                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                                Dim Name_of_the_block As String = TextBox_block_name.Text
                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then
                                        Dim Database1 As New Database(False, True)
                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try
                                        HostApplicationServices.WorkingDatabase = Database1
                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction


                                            'Database1.Aunits = UnitsValue.Meters


                                            Dim Layoutdict As DBDictionary = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            For Each entry As DBDictionaryEntry In Layoutdict
                                                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                For Each Id1 As ObjectId In BTrecord
                                                    Dim ent1 As Entity = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), Entity)
                                                    If IsNothing(ent1) = False Then
                                                        If TypeOf (ent1) Is MText = True Then
                                                            ent1.UpgradeOpen()
                                                            Dim Mtext1 As MText = ent1
                                                            If Mtext1.Contents = "⅊" Then
                                                                Dim NEW_TXT As String = "{\fISOCPEUR|b0|i0|c0|p34;" + Mtext1.Contents + "}"

                                                                Mtext1.Contents = NEW_TXT
                                                            End If

                                                        End If
                                                    End If
                                                Next
                                            Next

                                            Trans1.Commit()
                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using
                                        Database1.Dispose()
                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If
123:
                                Next
                                Trans11.Commit()
                            End Using
                        End Using
                    End If
                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf & _
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Catch ex As System.SystemException
                MsgBox("1: " & ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_rectangle_Click(sender As Object, e As EventArgs) Handles Button_read_rectangle.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Empty_array() As ObjectId
            Dim This_drawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = This_drawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try

                Data_table_Viewport_data = New System.Data.DataTable
                Data_table_Viewport_data.Columns.Add("X", GetType(Double))
                Data_table_Viewport_data.Columns.Add("Y", GetType(Double))
                Data_table_Viewport_data.Columns.Add("ROT", GetType(Double))








                Using lock1 As DocumentLock = This_drawing.LockDocument


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = This_drawing.TransactionManager.StartTransaction


                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select Template Rectangle"

                        Object_Prompt.SingleOnly = True

                        Rezultat1 = Editor1.GetSelection(Object_Prompt)


                        If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Freeze_operations = False
                            Editor1.SetImpliedSelection(Empty_array)
                            Exit Sub
                        End If






                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim Poly1 As Polyline
                            Poly1 = Ent1
                            If Poly1.NumberOfVertices >= 4 Then

                                Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

                                Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near

                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

                                Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Specify viewport top left side point:")

                                PP1.AllowNone = False
                                Result_point1 = Editor1.GetPoint(PP1)
                                If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If


                                Dim Result_point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Specify viewport top right side point:")

                                PP2.AllowNone = False
                                PP2.UseBasePoint = True
                                PP2.BasePoint = Result_point1.Value
                                Result_point2 = Editor1.GetPoint(PP2)
                                If Result_point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)


                                Dim Point1 As New Point3d(Result_point1.Value.X, Result_point1.Value.Y, 0)
                                Dim Point2 As New Point3d(Result_point2.Value.X, Result_point2.Value.Y, 0)

                                Dim L1 As Double = Point1.DistanceTo(Point2)
                                Dim Rot As Double = GET_Bearing_rad(Point1.X, Point1.Y, Point2.X, Point2.Y)

                                Dim PonP1 As Point3d = Poly1.GetClosestPointTo(Point1, Vector3d.ZAxis, False)
                                Dim PonP2 As Point3d = Poly1.GetClosestPointTo(Point2, Vector3d.ZAxis, False)
                                Dim param1 As Integer = Round(Poly1.GetParameterAtPoint(PonP1), 0)
                                Dim param2 As Integer = Round(Poly1.GetParameterAtPoint(PonP2), 0)

                                If param1 = param2 Then
                                    MsgBox("review your points")
                                    Freeze_operations = False
                                    Exit Sub
                                End If

                                If param1 > param2 Then
                                    Dim t As Integer
                                    t = param2
                                    param2 = param1
                                    param1 = t
                                End If

                                Dim H As Double = 0
                                Dim X As Double = 0
                                Dim y As Double = 0
                                If param1 = 0 Then
                                    H = Poly1.GetPointAtParameter(param2).DistanceTo(Poly1.GetPointAtParameter(param2 + 1))
                                    X = (Poly1.GetPointAtParameter(param1).x + Poly1.GetPointAtParameter(param2 + 1).x) / 2
                                    y = (Poly1.GetPointAtParameter(param1).y + Poly1.GetPointAtParameter(param2 + 1).y) / 2
                                Else
                                    H = Poly1.GetPointAtParameter(param1).DistanceTo(Poly1.GetPointAtParameter(param1 - 1))
                                    X = (Poly1.GetPointAtParameter(param1 - 1).x + Poly1.GetPointAtParameter(param2).x) / 2
                                    y = (Poly1.GetPointAtParameter(param1 - 1).y + Poly1.GetPointAtParameter(param2).y) / 2
                                End If


                                Data_table_Viewport_data.Rows.Add()
                                Data_table_Viewport_data.Rows(0).Item("X") = X
                                Data_table_Viewport_data.Rows(0).Item("Y") = y
                                Data_table_Viewport_data.Rows(0).Item("ROT") = Rot


                            End If




                        End If







                        Trans1.Commit()

                    End Using


1254:


                End Using

end1:

                If IsNothing(Data_table_Viewport_data) = False Then
                    If Data_table_Viewport_data.Rows.Count > 0 Then
                        Add_to_clipboard_Data_table(Data_table_Viewport_data)
                    End If
                End If

            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub


    Private Sub Button_adjust_viewport_Click(sender As Object, e As EventArgs) Handles Button_adjust_viewport.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try

                If Data_table_Viewport_data.Rows.Count > 0 Then
                    Dim Index_existent As Integer = 1
                    For i = 0 To Data_table_Viewport_data.Rows.Count - 1

                        If IsDBNull(Data_table_Viewport_data.Rows(i).Item("ROT")) = False And _
                            IsDBNull(Data_table_Viewport_data.Rows(i).Item("X")) = False And _
                            IsDBNull(Data_table_Viewport_data.Rows(i).Item("Y")) = False Then



                            Dim Anno_scale_name1 As String = ""
                            Dim DWG_units As Integer
                            Using lock2 As DocumentLock = ThisDrawing.LockDocument
                                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                                    Object_Prompt.MessageForAdding = vbLf & "Select Viewport"
                                    Object_Prompt.SingleOnly = True
                                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                        Editor1.WriteMessage(vbLf & "Command:")
                                        Freeze_operations = False
                                        Editor1.SetImpliedSelection(Empty_array)
                                        Exit Sub
                                    End If

                                    Dim vpId As ObjectId
                                    If TypeOf Trans1.GetObject(Rezultat1.Value(0).ObjectId, OpenMode.ForRead) Is Polyline Then
                                        vpId = LayoutManager.Current.GetNonRectangularViewportIdFromClipId(Rezultat1.Value(0).ObjectId)
                                        If vpId = Nothing Then
                                            Freeze_operations = False
                                            Exit Sub
                                        End If
                                    Else
                                        vpId = Rezultat1.Value(0).ObjectId
                                    End If


                                    Dim Viewport1 As Viewport = TryCast(Trans1.GetObject(vpId, OpenMode.ForWrite), Viewport)
                                    If IsNothing(Viewport1) = False Then
                                        Dim H1 As Double = Viewport1.Height
                                        Dim W1 As Double = Viewport1.Width
                                        Dim viewport2 As New Viewport
                                        viewport2.Height = Viewport1.Height
                                        viewport2.Width = Viewport1.Width
                                        viewport2.Layer = Viewport1.Layer

                                        viewport2.CenterPoint = Viewport1.CenterPoint
                                        viewport2.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
                                        viewport2.ViewTarget = New Point3d(Data_table_Viewport_data.Rows(i).Item("X"), Data_table_Viewport_data.Rows(i).Item("Y"), 0) ' asta e pozitia viewport in MODEL space
                                        viewport2.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin
                                        viewport2.TwistAngle = 2 * PI - Data_table_Viewport_data.Rows(i).Item("ROT") ' asta e PT TWIST
                                        viewport2.CustomScale = 1
                                        BTrecord.AppendEntity(viewport2)
                                        Trans1.AddNewlyCreatedDBObject(viewport2, True)

                                        viewport2.Locked = True
                                        viewport2.On = True
                                        Viewport1.Erase()
                                        Trans1.Commit()
                                    End If
                                End Using
                            End Using
                        Else
                            Dim DEBUG As String
                            DEBUG = "INVESTIGATE"
                        End If
                    Next



                End If



            Catch ex As Exception
                Freeze_operations = False
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_read_atr_to_excel_Click(sender As Object, e As EventArgs) Handles Button_read_atr_to_excel.Click
        If Freeze_operations = False Then
            Freeze_operations = True

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = Get_active_worksheet_from_Excel()
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try
                Try


                    Dim Error_list As String = ""

                    Dim Start1 As Integer = 1
                    If IsNumeric(TextBox_transfer_start_row.Text) = True Then
                        Start1 = Abs(CInt(TextBox_transfer_start_row.Text))
                    End If
                    Dim End1 As Integer = 1
                    If IsNumeric(TextBox_transfer_end_row.Text) = True Then
                        End1 = Abs(CInt(TextBox_transfer_end_row.Text))
                    End If
                    Dim row0 As Integer = 1
                    If IsNumeric(TextBox_dwg_row.Text) = True Then
                        row0 = Abs(CInt(TextBox_dwg_row.Text))
                    End If

                    Dim Column_atr_name As String
                    Column_atr_name = TextBox_col_atr_name.Text

                    Dim Column_atr_value_start As String
                    Column_atr_value_start = TextBox_col_atr_value_start.Text

                    Dim Column_atr_value_end As String
                    Column_atr_value_end = TextBox_col_atr_value_end.Text

                    Dim Column_Block As String
                    Column_Block = TextBox_col_block_name.Text


                    Dim DataTable_blocks As New System.Data.DataTable()
                    DataTable_blocks.Columns.Add("BLOCKNAME", GetType(String))
                    DataTable_blocks.Columns.Add("ATRNAME", GetType(String))
                    Dim Lista_blocks As New Specialized.StringCollection

                    For i = Start1 To End1
                        DataTable_blocks.Rows.Add()
                        Dim bl As String = W1.Range(Column_Block & i).Value2
                        If bl <> "" Then
                            DataTable_blocks.Rows(DataTable_blocks.Rows.Count - 1).Item("BLOCKNAME") = bl
                            DataTable_blocks.Rows(DataTable_blocks.Rows.Count - 1).Item("ATRNAME") = W1.Range(Column_atr_name & i).Value2
                            If Lista_blocks.Contains(bl) = False Then
                                Lista_blocks.Add(bl)
                            End If
                        End If

                    Next

                    Dim Lista_dwg As New Specialized.StringCollection
                    Dim Col_value_start As Integer = W1.Range(Column_atr_value_start & "1").Column
                    Dim Col_value_end As Integer = W1.Range(Column_atr_value_end & "1").Column

                    For i = Col_value_start To Col_value_end
                        Lista_dwg.Add(W1.Cells(row0, i).value2)
                    Next


  


                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = ThisDrawing.LockDocument

                            Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()






                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    If IO.File.Exists(Drawing1) = True Then
                                        Dim dwg_name As String = IO.Path.GetFileNameWithoutExtension(Drawing1)
                                        If Lista_dwg.Contains(dwg_name) = True Then
                                            Dim Database1 As New Database(False, True)

                                            Try
                                                Try
                                                    Try
                                                        Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                    Catch ex As Exception
                                                        MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                        Error_list = Error_list & vbCrLf & Drawing1
                                                        GoTo 123
                                                    End Try
                                                Catch ex As IO.IOException
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As System.SystemException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try

                                            HostApplicationServices.WorkingDatabase = Database1
                                            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                                                Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                                Dim Layoutdict As DBDictionary

                                                Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                                Dim Index_datatable As Integer = 0


                                                Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)

                                                For k = 0 To Lista_blocks.Count - 1
                                                    Dim Blockname1 As String = Lista_blocks(k)
                                                    If BlockTable1.Has(Blockname1) = True Then
                                                        For Each entry As DBDictionaryEntry In Layoutdict
                                                            Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                            If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) Then
                                                                Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForRead)
                                                                For Each id1 As ObjectId In BTrecord
                                                                    Dim ent1 As Entity = TryCast(Trans1.GetObject(id1, OpenMode.ForRead), Entity)
                                                                    If IsNothing(ent1) = False Then
                                                                        If TypeOf ent1 Is BlockReference Then
                                                                            Dim Block1 As BlockReference = ent1
                                                                            Dim Block_name As String = ""

                                                                            Dim BlockTrec As BlockTableRecord = Nothing
                                                                            If Block1.IsDynamicBlock = True Then
                                                                                BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                                                Block_name = BlockTrec.Name
                                                                            Else
                                                                                BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                                                Block_name = BlockTrec.Name
                                                                            End If

                                                                            If Block_name = Blockname1 Then

                                                                                If Block1.AttributeCollection.Count > 0 Then
                                                                                    For Each Atid As ObjectId In Block1.AttributeCollection
                                                                                        Dim Atr1 As AttributeReference = TryCast(Atid.GetObject(OpenMode.ForRead), AttributeReference)

                                                                                        If IsNothing(Atr1) = False Then
                                                                                            For n = 0 To DataTable_blocks.Rows.Count - 1
                                                                                                Dim Bl1 As String = DataTable_blocks.Rows(n).Item("BLOCKNAME")
                                                                                                If Bl1 = Blockname1 Then
                                                                                                    If Atr1.Tag = DataTable_blocks.Rows(n).Item("ATRNAME") Then

                                                                                                        Dim File_name As String = IO.Path.GetFileNameWithoutExtension(Drawing1)

                                                                                                        If DataTable_blocks.Columns.Contains(File_name) = False Then
                                                                                                            DataTable_blocks.Columns.Add(File_name, GetType(String))
                                                                                                        End If
                                                                                                        DataTable_blocks.Rows(n).Item(File_name) = Atr1.TextString

                                                                                                        Exit For
                                                                                                    End If
                                                                                                End If
                                                                                            Next
                                                                                        End If
                                                                                    Next




                                                                                End If


                                                                            End If


                                                                        End If


                                                                    End If


                                                                Next


                                                            End If
                                                        Next
                                                    End If
                                                Next




                                                Trans1.Commit()

                                                Try

                                                Catch ex As Exception
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                End Try
                                            End Using

                                            Database1.Dispose()

                                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database


                                        End If











                                    End If

123:
                                Next


                                If DataTable_blocks.Rows.Count > 0 Then
                                    For i = 0 To DataTable_blocks.Rows.Count - 1
                                        For j = 2 To DataTable_blocks.Columns.Count - 1

                                            Dim File_name As String = DataTable_blocks.Columns(j).ColumnName
                                            Dim Valoare As String = ""
                                            If IsDBNull(DataTable_blocks.Rows(i).Item(j)) = False Then
                                                Valoare = DataTable_blocks.Rows(i).Item(j)
                                            End If
                                            W1.Cells(Start1 + i, Col_value_start + j - 2).value2 = Valoare
                                        Next
                                    Next
                                End If


                                Trans11.Commit()
                                MsgBox("DONE")
                            End Using
                        End Using
                    End If
                    If Not Error_list = "" Then
                        MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf & _
                               "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                        My.Computer.Clipboard.SetText(Error_list)

                    End If


                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                Catch ex As Exception
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    MsgBox(ex.Message)
                End Try
            Catch EX As System.Runtime.InteropServices.COMException
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(EX.Message)
            End Try

            Freeze_operations = False
        End If

    End Sub

    Private Sub Button_write_multiple_dwgs_Click(sender As Object, e As EventArgs) Handles Button_write_multiple_dwgs.Click
        If Freeze_operations = False Then
            Freeze_operations = True


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Empty_array() As ObjectId

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Editor1.SetImpliedSelection(Empty_array)
            Try




                Dim Start1 As Integer = 1
                If IsNumeric(TextBox_transfer_row_start.Text) = True Then
                    Start1 = Abs(CInt(TextBox_transfer_row_start.Text))
                End If
                Dim End1 As Integer = 0
                If IsNumeric(TextBox_transfer_row_end.Text) = True Then
                    End1 = Abs(CInt(TextBox_transfer_row_end.Text))
                End If

                If End1 < Start1 Then

                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub

                End If

                Dim Error_list As String = ""


                Dim Column_start As String = TextBox_block_att_value_column_start.Text.ToUpper
                Dim Column_end As String = TextBox_block_att_value_column_end.Text.ToUpper
                Dim Column_atr_name As String = TextBox_block_att_name_column.Text.ToUpper
                Dim Column_blockname As String = TextBox_block_att_blockname.Text.ToUpper

                Dim Col_start As Integer = 0
                Dim Col_end As Integer = 0
                Dim Col_atr As Integer = 0
                Dim Col_bl As Integer = 0

                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                W1 = Nothing


                Try
                    W1 = Get_active_worksheet_from_Excel_with_error()
                    If IsNothing(W1) = False Then
                        Col_start = W1.Range(Column_start & "1").Column
                        Col_end = W1.Range(Column_end & "1").Column
                        Col_atr = W1.Range(Column_atr_name & "1").Column
                        Col_bl = W1.Range(Column_blockname & "1").Column
                    End If

                Catch ex As System.SystemException

                    Editor1.WriteMessage(vbLf & "Command:")
                    Freeze_operations = False
                    Exit Sub

                End Try

                Dim Row0 As Integer = 0
                If IsNumeric(TextBox_ROW_layout_NAME.Text) = True Then
                    Row0 = CInt(TextBox_ROW_layout_NAME.Text)
                End If
                Dim Lista_dwg As New Specialized.StringCollection

                Dim Dt_block_info As New System.Data.DataTable()
                Dt_block_info.Columns.Add("BLOCKNAME", GetType(String))
                Dt_block_info.Columns.Add("ATRNAME", GetType(String))
                For i = Col_start To Col_end
                    Dim FileName1 As String = W1.Cells(Row0, i).value2
                    Dt_block_info.Columns.Add(FileName1, GetType(String))
                    Lista_dwg.Add(FileName1)
                Next

                Dim Lista_blocks As New Specialized.StringCollection


                For i = Start1 To End1
                    Dt_block_info.Rows.Add()
                    Dim Blname As String = W1.Cells(i, Col_bl).VALUE2
                    Dt_block_info.Rows(Dt_block_info.Rows.Count - 1).Item("BLOCKNAME") = Blname
                    Dt_block_info.Rows(Dt_block_info.Rows.Count - 1).Item("ATRNAME") = W1.Cells(i, Col_atr).VALUE2

                    If Lista_blocks.Contains(Blname) = False Then
                        Lista_blocks.Add(Blname)
                    End If

                    Dim Index_col As Integer = 2

                    For j = Col_start To Col_end
                        Dim Val1 As String = W1.Cells(i, j).VALUE2
                        If Val1 <> "" Then
                            Dt_block_info.Rows(Dt_block_info.Rows.Count - 1).Item(Index_col) = Val1
                        End If
                        Index_col = Index_col + 1
                    Next
                Next

                Add_to_clipboard_Data_table(Dt_block_info)

                If ListBox_DWG.Items.Count > 0 Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument

                        Using Trans11 As Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                            For i = 0 To ListBox_DWG.Items.Count - 1
                                Dim Drawing1 As String = ListBox_DWG.Items(i)
                                If IO.File.Exists(Drawing1) = True Then
                                    Dim dwg_name As String = IO.Path.GetFileNameWithoutExtension(Drawing1)
                                    Dim index1 As Integer = -1

                                    For j = 2 To Dt_block_info.Columns.Count - 1
                                        Dim File_name As String = Dt_block_info.Columns(j).ColumnName
                                        If File_name = dwg_name Then
                                            index1 = j
                                            Exit For
                                        End If
                                    Next

                                    If Lista_dwg.Contains(dwg_name) = True And index1 > -1 Then
                                        Dim Database1 As New Database(False, True)
                                        Try
                                            Try
                                                Try
                                                    Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                                Catch ex As Exception
                                                    MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                    Error_list = Error_list & vbCrLf & Drawing1
                                                    GoTo 123
                                                End Try
                                            Catch ex As IO.IOException
                                                MsgBox(Drawing1 & vbCrLf & "could not be open")
                                                Error_list = Error_list & vbCrLf & Drawing1
                                                GoTo 123
                                            End Try
                                        Catch ex As System.SystemException
                                            MsgBox(Drawing1 & vbCrLf & "could not be open")
                                            Error_list = Error_list & vbCrLf & Drawing1
                                            GoTo 123
                                        End Try
                                        HostApplicationServices.WorkingDatabase = Database1
                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                                            Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                                            Dim Layoutdict As DBDictionary
                                            Layoutdict = Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead)
                                            Dim Index_datatable As Integer = 0
                                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                                            For k = 0 To Lista_blocks.Count - 1
                                                Dim Blockname1 As String = Lista_blocks(k)
                                                If BlockTable1.Has(Blockname1) = True Then
                                                    For Each entry As DBDictionaryEntry In Layoutdict
                                                        Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite)
                                                        If Layout1.TabOrder >= CInt(TextBox_Layout_index_START.Text) And Layout1.TabOrder <= CInt(TextBox_Layout_index_END.Text) Then
                                                            If Layoutdict.Count = 2 Then
                                                                Layout1.LayoutName = IO.Path.GetFileNameWithoutExtension(Drawing1).ToString
                                                            End If

                                                            Dim BTrecord As BlockTableRecord = Trans1.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite)
                                                            For Each id1 As ObjectId In BTrecord
                                                                Dim ent1 As Entity = TryCast(Trans1.GetObject(id1, OpenMode.ForRead), Entity)
                                                                If IsNothing(ent1) = False Then
                                                                    If TypeOf ent1 Is BlockReference Then
                                                                        Dim Block1 As BlockReference = ent1
                                                                        Dim Block_name As String = ""

                                                                        Dim BlockTrec As BlockTableRecord = Nothing
                                                                        If Block1.IsDynamicBlock = True Then
                                                                            BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, OpenMode.ForRead)
                                                                            Block_name = BlockTrec.Name
                                                                        Else
                                                                            BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, OpenMode.ForRead)
                                                                            Block_name = BlockTrec.Name
                                                                        End If

                                                                        If Block_name = Blockname1 Then

                                                                            If Block1.AttributeCollection.Count > 0 Then
                                                                                Block1.UpgradeOpen()

                                                                                For Each Atid As ObjectId In Block1.AttributeCollection
                                                                                    Dim Atr1 As AttributeReference = TryCast(Atid.GetObject(OpenMode.ForWrite), AttributeReference)



                                                                                    If IsNothing(Atr1) = False Then
                                                                                        Dim tag1 As String = Atr1.Tag

                                                                                        For n = 0 To Dt_block_info.Rows.Count - 1

                                                                                            Dim Bl1 As String = Dt_block_info.Rows(n).Item("BLOCKNAME")
                                                                                            If Bl1 = Blockname1 Then

                                                                                                If tag1 = Dt_block_info.Rows(n).Item("ATRNAME") Then
                                                                                                    Dim Valoare = ""
                                                                                                    If IsDBNull(Dt_block_info.Rows(n).Item(index1)) = False Then
                                                                                                        Valoare = Dt_block_info.Rows(n).Item(index1)
                                                                                                    End If
                                                                                                    Atr1.TextString = Valoare

                                                                                                    If Atr1.IsMTextAttribute = False Then

                                                                                                    Else
                                                                                                        'Atr1.MTextAttribute.Contents = Valoare
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Next
                                                                                    End If
                                                                                Next
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                    Next
                                                End If
                                            Next
                                            Trans1.Commit()

                                            Try
                                                Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                            Catch ex As Exception
                                                Error_list = Error_list & vbCrLf & Drawing1
                                            End Try
                                        End Using
                                        Database1.Dispose()
                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database
                                    End If
                                End If

123:
                            Next

                            Trans11.Commit()
                            MsgBox("DONE")
                        End Using
                    End Using
                End If
                If Not Error_list = "" Then
                    MsgBox("There are not saved files. The list of these files is loaded into the clipboard" & vbCrLf & _
                           "Use paste <Ctrl+V> command in a text editor program (notepad, excel, word) to see the list")
                    My.Computer.Clipboard.SetText(Error_list)

                End If

                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try

            Freeze_operations = False
        End If

        

    End Sub



End Class