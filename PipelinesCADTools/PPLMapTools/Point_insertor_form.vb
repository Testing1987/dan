Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Point_insertor_form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Col_east2 As Integer
    Dim Col_north2 As Integer

    Private Sub insertor_form_Load(sender As Object, e As System.EventArgs) Handles Me.Load


        Incarca_existing_layers_to_combobox(ComboBox_Layer_for_blocks)
        ComboBox_Layer_for_blocks.SelectedIndex = 0
        If ComboBox_Layer_for_blocks.Items.Contains("TEXT") = True Then
            ComboBox_Layer_for_blocks.SelectedIndex = ComboBox_Layer_for_blocks.Items.IndexOf("TEXT")
        End If
        If ComboBox_Layer_for_blocks.Items.Contains("Text") = True Then
            ComboBox_Layer_for_blocks.SelectedIndex = ComboBox_Layer_for_blocks.Items.IndexOf("Text")
        End If
        If ComboBox_Layer_for_blocks.Items.Contains("text") = True Then
            ComboBox_Layer_for_blocks.SelectedIndex = ComboBox_Layer_for_blocks.Items.IndexOf("text")
        End If
        With TextBox_Point_name
            .Select()
        End With


        Panel_BLOCKS.Visible = False

        Incarca_existing_layers_to_combobox(ComboBox_poly_layer)
        ComboBox_poly_layer.SelectedIndex = 0
    End Sub

    Private Sub CheckBox_insert_blocks_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox_insert_blocks.CheckedChanged
        If CheckBox_insert_blocks.Checked = True Then
            Panel_BLOCKS.Visible = True

        Else
            Panel_BLOCKS.Visible = False

        End If
    End Sub

    Private Sub Button_2D_3D_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_2D_3D.Click
        If Button_2D_3D.Text = "2D" Then
            Button_2D_3D.Text = "3D"
            Button_2D_3D.BackColor = Drawing.Color.Magenta
        Else
            Button_2D_3D.Text = "2D"
            Button_2D_3D.BackColor = Drawing.Color.Indigo
        End If
    End Sub

    Private Sub Button_remove_points_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_remove_points.Click
        Try
            If MsgBox("Are you sure you want to remove all survey data?", MsgBoxStyle.YesNo, "Insertor") = MsgBoxResult.Yes Then



                TextBox_message.Text = "Work in progress ..Don't move..Be patient"
                Me.Refresh()
                Dim Lock1 As DocumentLock
                Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument


                Using Lock1
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
                    Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database
                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

                    ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Database1 = ThisDrawing.Database
                    Editor1 = ThisDrawing.Editor


                    Dim Lista_layere_data As New Specialized.StringCollection

                    Dim Filtru1(0) As TypedValue

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                        LayerTable1 = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim Layer0_id As ObjectId = LayerTable1.Item("0")
                        Database1.Clayer = Layer0_id

                        For Each Obj_Id1 As ObjectId In LayerTable1
                            Dim Layer1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                            Layer1 = Trans1.GetObject(Obj_Id1, OpenMode.ForWrite)
                            If Layer1.Description = "created by point insertor" And Layer1.IsPlottable = False Then
                                Filtru1(0) = New Autodesk.AutoCAD.DatabaseServices.TypedValue(Autodesk.AutoCAD.DatabaseServices.DxfCode.LayerName,
                                                                                              Layer1.Name)
                                Dim Selection_Filter1 As New Autodesk.AutoCAD.EditorInput.SelectionFilter(Filtru1)

                                Dim Selection_result1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                                Selection_result1 = Editor1.SelectAll(Selection_Filter1)

                                Dim Selset1 As Autodesk.AutoCAD.EditorInput.SelectionSet
                                Selset1 = Selection_result1.Value
                                If Not Selset1 Is Nothing Then
                                    If Selset1.Count > 0 Then
                                        For i = 0 To Selset1.Count - 1
                                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                            Obj1 = Selset1.Item(i)
                                            Dim Ent1 As Autodesk.AutoCAD.DatabaseServices.Entity
                                            Ent1 = Trans1.GetObject(Obj1.ObjectId, OpenMode.ForWrite)
                                            Ent1.Erase()
                                        Next
                                    End If
                                End If

                                Layer1.Erase()
                            End If
                        Next

                        Trans1.Commit()
                        ' asta e de la tranzactie
                    End Using
                    ' asta e de la lock
                End Using

                TextBox_message.Text = "The survey data has been removed"

                ' asta e de la mesage box
            Else
                TextBox_message.Text = "You just cancel survey data removal"
                ' asta e de la mesage box
            End If

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Public Overridable Function InsertBlock(ByVal dblInsert As Point3d, ByVal btrSpace As BlockTableRecord, ByVal strSourceBlockName As String,
                                             ByVal Layer1 As String, ByVal Scale1 As Double) As BlockReference
        Dim dlock As DocumentLock = Nothing
        Dim bt As BlockTable
        Dim btr As BlockTableRecord = Nothing
        Dim br As BlockReference
        Dim id As ObjectId
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor

            'insert block and rename it
            Try
                Try
                    dlock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Catch ex As Exception
                    Dim aex As New System.Exception("Error locking document for InsertBlock: " & strSourceBlockName & ": ", ex)
                    Throw aex
                End Try
                bt = trans.GetObject(db.BlockTableId, OpenMode.ForWrite)
                If bt.Has(strSourceBlockName) Then
                    'block found, get instance for copying
                    btr = trans.GetObject(bt.Item(strSourceBlockName), OpenMode.ForRead)
                Else
                    Return Nothing
                    Exit Function
                    'CloneBlock(strSourceBlockPath, "*ModelSpace") ''clone block will not work when inserting the entire 'current drawing'
                End If
                'If bt.Has(strSourceBlockName) Then MsgBox("Got it: " & strSourceBlockName)
                btrSpace = trans.GetObject(btrSpace.ObjectId, OpenMode.ForWrite)
                'Set the Attribute Value
                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                Dim ent As Entity
                Dim btrenum As BlockTableRecordEnumerator
                br = New BlockReference(dblInsert, btr.ObjectId)
                br.Layer = Layer1
                br.ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(Scale1, Scale1, Scale1)

                btrSpace.AppendEntity(br)
                trans.AddNewlyCreatedDBObject(br, True)
                attColl = br.AttributeCollection
                btrenum = btr.GetEnumerator
                While btrenum.MoveNext
                    ent = btrenum.Current.GetObject(OpenMode.ForWrite)
                    If TypeOf ent Is AttributeDefinition Then
                        Dim attdef As AttributeDefinition = ent
                        Dim attref As New AttributeReference
                        attref.SetAttributeFromBlock(attdef, br.BlockTransform)
                        attref.TextString = attref.Tag
                        attColl.AppendAttribute(attref)
                        trans.AddNewlyCreatedDBObject(attref, True)
                    End If
                End While
                trans.Commit()
            Catch ex As System.Exception
                Dim aex2 As New System.Exception("Error in inserting new block: " & strSourceBlockName & ": ", ex)
                Throw aex2
            Finally
                If Not trans Is Nothing Then trans.Dispose()
                If Not dlock Is Nothing Then dlock.Dispose()
            End Try
        End Using
        Return br
    End Function

    Public Function Get_String_Rounded(ByVal Numar As Double, ByVal Nr_dec As Integer) As String
        Dim String1, String2, Zero, zero1 As String
        Zero = ""
        zero1 = ""

        Dim String_punct As String = ""

        If Nr_dec > 0 Then
            String_punct = "."
            For i = 1 To Nr_dec
                Zero = Zero & "0"
            Next
        End If

        Dim String_minus As String = ""

        If Numar < 0 Then
            String_minus = "-"
            Numar = -Numar
        End If

        String1 = Round(Numar, Nr_dec).ToString
        String2 = String1

        If InStr(String1, ".") = 0 Then
            String2 = String1 & String_punct & Zero
            GoTo 123
        End If

        If Len(String1) - InStr(String1, ".") - Nr_dec <> 0 Then
            For i = 1 To InStr(String1, ".") + Nr_dec - Len(String1)
                zero1 = zero1 & "0"
            Next

            String2 = String1 & zero1
        End If
123:
        Return String_minus & String2
    End Function


    Public Function Stabileste_litera_coloana(ByVal nr1 As Integer) As String
        Select Case UCase(nr1)
            Case (1)
                Return "A"
            Case (2)
                Return "B"
            Case (3)
                Return "C"
            Case (4)
                Return "D"
            Case (5)
                Return "E"
            Case (6)
                Return "F"
            Case (7)
                Return "G"
            Case (8)
                Return "H"
            Case (9)
                Return "I"
            Case (10)
                Return "J"
            Case (11)
                Return "K"
            Case (12)
                Return "L"
            Case (13)
                Return "M"
            Case (14)
                Return "N"
            Case (15)
                Return "O"
            Case (16)
                Return "P"
            Case (17)
                Return "Q"
            Case (18)
                Return "R"
            Case (19)
                Return "S"
            Case (20)
                Return "T"
            Case (21)
                Return "U"
            Case (22)
                Return "V"
            Case (23)
                Return "W"
            Case (24)
                Return "X"
            Case (25)
                Return "Y"
            Case (26)
                Return "Z"
            Case (27)
                Return "AA"
            Case (28)
                Return "AB"
            Case (29)
                Return "AC"
            Case (30)
                Return "AD"
            Case (31)
                Return "AE"
            Case (32)
                Return "AF"
            Case (33)
                Return "AG"
            Case (34)
                Return "AH"
            Case (35)
                Return "AI"
            Case (36)
                Return "AJ"
            Case (37)
                Return "AK"
            Case (38)
                Return "AL"
            Case (39)
                Return "AM"
            Case (40)
                Return "AN"
            Case (41)
                Return "AO"
            Case (42)
                Return "AP"
            Case (43)
                Return "AQ"
            Case (44)
                Return "AR"
            Case (45)
                Return "AS"
            Case (46)
                Return "AT"
            Case (47)
                Return "AU"
            Case (48)
                Return "AV"
            Case (49)
                Return "AW"
            Case (50)
                Return "AX"
            Case (51)
                Return "AY"
            Case (52)
                Return "AZ"
        End Select
    End Function


    Public Function Load_Excel_to_data_table(ByVal Start1 As Double, ByVal End1 As Double, ByVal Col_PN As Integer, ByVal Col_East As Integer, ByVal Col_North As Integer, ByVal Col_Elevation As Integer,
                                             ByVal Col_Code As Integer, ByVal Col_Lcode As Integer, ByVal Col_Descriptie_Custom As Integer, ByVal Col_Block_name As Integer,
                                             ByVal Col_Atr_tag1 As Integer, ByVal Col_Atr_value1 As Integer, ByVal Col_Atr_tag2 As Integer, ByVal Col_Atr_value2 As Integer, ByVal Col_ln As Integer) As System.Data.DataTable

        Try

            Dim Table_data1 As New System.Data.DataTable
            Table_data1.Columns.Add("PN", GetType(String))
            Table_data1.Columns.Add("East", GetType(String))
            Table_data1.Columns.Add("North", GetType(String))
            Table_data1.Columns.Add("Elevation", GetType(String))
            Table_data1.Columns.Add("Code", GetType(String))
            Table_data1.Columns.Add("LCode", GetType(String))
            Table_data1.Columns.Add("Descriptie_custom", GetType(String))
            Table_data1.Columns.Add("Block_name", GetType(String))
            Table_data1.Columns.Add("Attribute_tag1", GetType(String))
            Table_data1.Columns.Add("Attribute_value1", GetType(String))
            Table_data1.Columns.Add("Attribute_tag2", GetType(String))
            Table_data1.Columns.Add("Attribute_value2", GetType(String))
            Table_data1.Columns.Add("layer", GetType(String))


            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet

            W1 = Get_active_worksheet_from_Excel()


            Dim Index_data_table As Integer = 0

            For i = Start1 To End1

                Table_data1.Rows.Add()

                If Col_PN > 0 Then
                    Table_data1.Rows.Item(Index_data_table).Item("PN") = W1.Cells(i, Col_PN).value2
                End If

                If Col_East > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_East).value2

                    If IsNumeric(Cell_value) = True Then
                        Table_data1.Rows.Item(Index_data_table).Item("East") = Cell_value
                    End If
                End If

                If Col_North > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_North).value2

                    If IsNumeric(Cell_value) = True Then
                        Table_data1.Rows.Item(Index_data_table).Item("North") = Cell_value
                    End If
                End If

                If Col_Elevation > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Elevation).value2


                    Table_data1.Rows.Item(Index_data_table).Item("Elevation") = Cell_value

                Else
                    Table_data1.Rows.Item(Index_data_table).Item("Elevation") = "0"
                End If


                If Col_Code > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Code).value2
                    If IsNumeric(Cell_value) = True Then
                        If CInt(Cell_value) < 1000 And CInt(Cell_value) > 0 Then
                            Table_data1.Rows.Item(Index_data_table).Item("Code") = CInt(Cell_value)
                        End If
                    End If
                End If


                If Col_Lcode > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Lcode).value2
                    If Cell_value = "1" Or Cell_value = "2" Or Cell_value = "3" Then
                        Table_data1.Rows.Item(Index_data_table).Item("LCode") = Cell_value
                    Else
                        Table_data1.Rows.Item(Index_data_table).Item("LCode") = "0"
                    End If

                Else
                    Table_data1.Rows.Item(Index_data_table).Item("LCode") = "0"
                End If



                If Col_Descriptie_Custom > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Descriptie_Custom).value2
                    Table_data1.Rows.Item(Index_data_table).Item("Descriptie_custom") = Cell_value
                End If





                If Col_Block_name > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Block_name).value2
                    Table_data1.Rows.Item(Index_data_table).Item("Block_name") = Cell_value
                End If



                If Col_Atr_tag1 > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Atr_tag1).value2
                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_tag1") = Cell_value
                End If



                If Col_Atr_value1 > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Atr_value1).value2
                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_value1") = Cell_value
                End If




                If Col_Atr_tag2 > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Atr_tag2).value2
                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_tag2") = Cell_value
                End If



                If Col_Atr_value2 > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_Atr_value2).value2
                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_value2") = Cell_value
                End If

                If Col_ln > 0 Then
                    Dim Cell_value As String = W1.Cells(i, Col_ln).value2
                    Table_data1.Rows.Item(Index_data_table).Item("layer") = Cell_value
                End If


                Index_data_table = Index_data_table + 1

            Next

            Return Table_data1


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function


    Public Function Load_Excel_prin_copy_paste_to_data_table(ByVal Start1 As Double, ByVal End1 As Double, ByVal Col_PN As Integer, ByVal Col_East As Integer, ByVal Col_North As Integer, ByVal Col_Elevation As Integer,
                                             ByVal Col_Code As Integer, ByVal Col_Lcode As Integer, ByVal Col_Descriptie_Custom As Integer, ByVal Col_Block_name As Integer,
                                             ByVal Col_Atr_tag1 As Integer, ByVal Col_Atr_value1 As Integer, ByVal Col_Atr_tag2 As Integer, ByVal Col_Atr_value2 As Integer) As System.Data.DataTable

        Try

            Dim Table_data1 As New System.Data.DataTable
            Table_data1.Columns.Add("PN", GetType(String))
            Table_data1.Columns.Add("East", GetType(String))
            Table_data1.Columns.Add("North", GetType(String))
            Table_data1.Columns.Add("Elevation", GetType(String))
            Table_data1.Columns.Add("Code", GetType(String))
            Table_data1.Columns.Add("LCode", GetType(String))
            Table_data1.Columns.Add("Descriptie_custom", GetType(String))
            Table_data1.Columns.Add("Block_name", GetType(String))
            Table_data1.Columns.Add("Attribute_tag1", GetType(String))
            Table_data1.Columns.Add("Attribute_value1", GetType(String))
            Table_data1.Columns.Add("Attribute_tag2", GetType(String))
            Table_data1.Columns.Add("Attribute_value2", GetType(String))
            Dim Path_to_desktop As String = Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)
            Dim Nume_fisier_csv As String = Path_to_desktop & "\points.csv"


            Dim Excel1 As Microsoft.Office.Interop.Excel.Application
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)

            Try
                Excel1.DisplayAlerts = False
            Catch ex As Runtime.InteropServices.COMException
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Function
            End Try

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet

            W1 = Get_active_worksheet_from_Excel()

            W1.Rows(Start1 & ":" & End1).Copy()
            Excel1.Workbooks.Add()
            Dim W2 As Microsoft.Office.Interop.Excel.Worksheet

            W2 = Get_active_worksheet_from_Excel()
            W2.Range("A1").Select()
            W2.Range("A1").PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Operation:=Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
            W2.Cells.MergeCells = False


            W2.SaveAs(Nume_fisier_csv, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV)
            Excel1.ActiveWindow.Close()

            If System.IO.File.Exists(Nume_fisier_csv) = True Then

                Using Reader1 As New System.IO.StreamReader(Nume_fisier_csv)
                    Dim Line1 As String
                    Dim Index_data_table As Integer = 0

                    While Reader1.Peek > 0
                        Line1 = Reader1.ReadLine
                        Table_data1.Rows.Add()
                        If InStr(Line1, ",") > 0 Then
                            Dim Bucati_linie() As String
                            Bucati_linie = Split(Line1, ",")

                            If Col_PN > 0 And Bucati_linie.Length >= Col_PN Then
                                Table_data1.Rows.Item(Index_data_table).Item("PN") = Bucati_linie(Col_PN - 1)
                            End If

                            If Col_East > 0 And Bucati_linie.Length >= Col_East Then
                                If IsNumeric(Bucati_linie(Col_East - 1)) = True And Bucati_linie.Length > Col_East Then
                                    Table_data1.Rows.Item(Index_data_table).Item("East") = Bucati_linie(Col_East - 1)
                                End If
                            End If

                            If Col_North > 0 And Bucati_linie.Length >= Col_North Then
                                If IsNumeric(Bucati_linie(Col_North - 1)) = True Then
                                    Table_data1.Rows.Item(Index_data_table).Item("North") = Bucati_linie(Col_North - 1)
                                End If
                            End If

                            If Col_Elevation > 0 And Bucati_linie.Length >= Col_Elevation Then
                                If IsNumeric(Bucati_linie(Col_Elevation - 1)) = True Then
                                    Table_data1.Rows.Item(Index_data_table).Item("Elevation") = Bucati_linie(Col_Elevation - 1)
                                Else
                                    Table_data1.Rows.Item(Index_data_table).Item("Elevation") = "0"
                                End If
                            Else
                                Table_data1.Rows.Item(Index_data_table).Item("Elevation") = "0"
                            End If


                            If Col_Code > 0 And Bucati_linie.Length >= Col_Code Then
                                If IsNumeric(Bucati_linie(Col_Code - 1)) = True Then
                                    If CInt(Bucati_linie(Col_Code - 1)) < 1000 And CInt(Bucati_linie(Col_Code - 1)) > 0 Then
                                        Table_data1.Rows.Item(Index_data_table).Item("Code") = CInt(Bucati_linie(Col_Code - 1))
                                    End If
                                End If
                            End If

                            If Bucati_linie.Length >= Col_Lcode Then
                                If Col_Lcode > 0 Then

                                    If Bucati_linie(Col_Lcode - 1) = "1" Or Bucati_linie(Col_Lcode - 1) = "2" Or Bucati_linie(Col_Lcode - 1) = "3" Then
                                        Table_data1.Rows.Item(Index_data_table).Item("LCode") = Bucati_linie(Col_Lcode - 1)
                                    Else
                                        Table_data1.Rows.Item(Index_data_table).Item("LCode") = "0"
                                    End If

                                Else
                                    Table_data1.Rows.Item(Index_data_table).Item("LCode") = "0"
                                End If
                            Else
                                Table_data1.Rows.Item(Index_data_table).Item("LCode") = "0"
                            End If

                            If Bucati_linie.Length >= Col_Descriptie_Custom Then
                                If Col_Descriptie_Custom > 0 Then
                                    Table_data1.Rows.Item(Index_data_table).Item("Descriptie_custom") = Bucati_linie(Col_Descriptie_Custom - 1)
                                End If
                            End If



                            If Bucati_linie.Length >= Col_Block_name Then
                                If Col_Block_name > 0 Then
                                    Table_data1.Rows.Item(Index_data_table).Item("Block_name") = Bucati_linie(Col_Block_name - 1)
                                End If
                            End If

                            If Bucati_linie.Length >= Col_Atr_tag1 Then
                                If Col_Atr_tag1 > 0 Then
                                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_tag1") = Bucati_linie(Col_Atr_tag1 - 1)
                                End If
                            End If

                            If Bucati_linie.Length >= Col_Atr_value1 Then
                                If Col_Atr_value1 > 0 Then
                                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_value1") = Bucati_linie(Col_Atr_value1 - 1)
                                End If
                            End If


                            If Bucati_linie.Length >= Col_Atr_tag2 Then
                                If Col_Atr_tag2 > 0 Then
                                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_tag2") = Bucati_linie(Col_Atr_tag2 - 1)
                                End If
                            End If

                            If Bucati_linie.Length >= Col_Atr_value2 Then
                                If Col_Atr_value2 > 0 Then
                                    Table_data1.Rows.Item(Index_data_table).Item("Attribute_value2") = Bucati_linie(Col_Atr_value2 - 1)
                                End If
                            End If

                            Index_data_table = Index_data_table + 1


                        End If


                        'asta e de la reader.peek >0
                    End While

                    'asta e de la reader
                End Using

                System.IO.File.Delete(Nume_fisier_csv)
                Excel1.DisplayAlerts = True
                'asta e de la fisierul exista
                Return Table_data1
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Private Sub Button_insert_points_to_acad_Click(sender As System.Object, e As System.EventArgs) Handles Button_insert_points_to_acad.Click
        Try

            TextBox_message.Text = "Work in progress ..Don't move..Be patient"

            If Val(TextBox_row_start.Text) < 1 Or IsNumeric(TextBox_row_start.Text) = False Then
                With TextBox_row_start
                    .Text = ""
                    .Focus()
                End With
                TextBox_message.Text = "Please specify start row"

                Exit Sub
            End If

            If Val(TextBox_row_end.Text) < 1 Or IsNumeric(TextBox_row_end.Text) = False Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                TextBox_message.Text = "Please specify end row"

                Exit Sub
            End If

            If Val(TextBox_row_end.Text) < Val(TextBox_row_start.Text) Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                TextBox_message.Text = "End row smaller than start row"

                Exit Sub
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Me.Refresh()
            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                Dim Col_PN1 As Integer = Stabileste_coloanele(TextBox_Point_name.Text)
                Dim Col_East1 As Integer = Stabileste_coloanele(TextBox_East.Text)
                Dim Col_North1 As Integer = Stabileste_coloanele(TextBox_NORTH.Text)
                Dim Col_Elevation1 As Integer = Stabileste_coloanele(TextBox_elevation.Text)
                Dim Col_Code1 As Integer = Stabileste_coloanele(TextBox_extra1.Text)
                Dim Col_Lcode1 As Integer = Stabileste_coloanele(TextBox_extra2.Text)
                Dim Col_Descriptie_Custom1 As Integer = Stabileste_coloanele(TextBox_description.Text)
                Dim Col_Block_name1 As Integer = Stabileste_coloanele(TextBox_block_name.Text)
                Dim Col_Atr_tag1 As Integer = Stabileste_coloanele(TextBox_Atribut_name1.Text)
                Dim Col_Atr_val1 As Integer = Stabileste_coloanele(TextBox_atribut_value1.Text)
                Dim Col_Atr_tag2 As Integer = Stabileste_coloanele(TextBox_Atribut_name2.Text)
                Dim Col_Atr_val2 As Integer = Stabileste_coloanele(TextBox_atribut_value2.Text)
                Dim Col_ln As Integer = Stabileste_coloanele(TextBox_ln.Text)

                Dim Table_data1 As New System.Data.DataTable

                Table_data1 = Load_Excel_to_data_table(Val(TextBox_row_start.Text), Val(TextBox_row_end.Text), Col_PN1, Col_East1, Col_North1, Col_Elevation1, Col_Code1, Col_Lcode1, Col_Descriptie_Custom1,
                                                                       Col_Block_name1, Col_Atr_tag1, Col_Atr_val1, Col_Atr_tag2, Col_Atr_val2, Col_ln)


                Dim Start_index_line_code As New DoubleCollection
                Dim Layer_line_code As New Specialized.StringCollection
                Dim lista_coduri As New IntegerCollection
                Dim Nr_de_linii As Integer = 0
                ' here is the checking part
                Dim Is3D As Boolean
                If Button_2D_3D.Text = "3D" Then
                    Is3D = True
                Else
                    Is3D = False
                End If

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Double = CDbl(TextBox_row_start.Text)
                    Dim end1 As Double = CDbl(TextBox_row_end.Text)



                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                    For i = 0 To Table_data1.Rows.Count - 1
                        Dim East1, North1, Elev1 As Double
                        Dim PN1 As String
                        Dim elev_string As String = ""
                        ' de aici

                        If IsDBNull(Table_data1.Rows.Item(i).Item("PN")) = False Then
                            PN1 = Table_data1.Rows.Item(i).Item("PN").ToString
                        End If


                        If IsDBNull(Table_data1.Rows.Item(i).Item("East")) = False Then
                            If IsNumeric(Table_data1.Rows.Item(i).Item("East")) = True Then
                                East1 = CDbl(Convert.ToString(Table_data1.Rows.Item(i).Item("East")).Replace("+", "").Replace("'", "").Replace(" ", ""))
                            Else
                                W1.Cells(i + start1, Col_east2).Interior.ColorIndex = 6
                                W1.Cells(i + start1, Col_east2).Select()
                                TextBox_message.Text = "See yellow cell - column " & Stabileste_litera_coloana(Col_east2) & ", row " & (i + start1) & " East"
                                GoTo 123
                            End If
                        Else
                            W1.Cells(i + start1, Col_east2).Interior.ColorIndex = 6
                            W1.Cells(i + start1, Col_east2).Select()
                            TextBox_message.Text = "See yellow cell - column " & Stabileste_litera_coloana(Col_east2) & ", row " & (i + start1) & " East"
                            GoTo 123
                        End If


                        If IsDBNull(Table_data1.Rows.Item(i).Item("North")) = False Then
                            If IsNumeric(Table_data1.Rows.Item(i).Item("North")) = True Then
                                North1 = CDbl(Convert.ToString(Table_data1.Rows.Item(i).Item("North")).Replace("+", "").Replace("'", "").Replace(" ", ""))
                            Else
                                W1.Cells(i + start1, Col_north2).Interior.ColorIndex = 6
                                W1.Cells(i + start1, Col_north2).Select()
                                TextBox_message.Text = "See yellow cell - column " & Stabileste_litera_coloana(Col_north2) & ", row " & (i + start1) & " North"
                                GoTo 123
                            End If
                        Else
                            W1.Cells(i + start1, Col_north2).Interior.ColorIndex = 6
                            W1.Cells(i + start1, Col_north2).Select()
                            TextBox_message.Text = "See yellow cell - column " & Stabileste_litera_coloana(Col_north2) & ", row " & (i + start1) & " North"
                            GoTo 123
                        End If

                        If IsDBNull(Table_data1.Rows.Item(i).Item("Elevation")) = False Then

                            If IsNumeric(Table_data1.Rows.Item(i).Item("Elevation")) = True Then
                                Elev1 = CDbl(Convert.ToString(Table_data1.Rows.Item(i).Item("Elevation").Replace("+", "").Replace("'", "").Replace(" ", "")))
                            Else
                                Elev1 = 0
                                elev_string = Convert.ToString(Table_data1.Rows.Item(i).Item("Elevation"))
                            End If

                        End If


                        Dim Code1 As Integer

                        If IsDBNull(Table_data1.Rows.Item(i).Item("Code")) = False Then
                            If IsNumeric(Table_data1.Rows.Item(i).Item("Code")) = True Then
                                Code1 = CInt(Table_data1.Rows.Item(i).Item("Code"))
                            Else
                                Code1 = 0
                            End If
                        Else
                            Code1 = 0
                        End If

                        Dim Lcode1 As Integer

                        If IsDBNull(Table_data1.Rows.Item(i).Item("LCode")) = False Then
                            If IsNumeric(Table_data1.Rows.Item(i).Item("LCode").ToString) = True Then
                                Lcode1 = Table_data1.Rows.Item(i).Item("LCode")
                            Else
                                Lcode1 = 0
                            End If
                        Else
                            Lcode1 = 0
                        End If

                        Dim Descriptie1 As String = ""

                        If IsDBNull(Table_data1.Rows.Item(i).Item("Descriptie_custom")) = False Then
                            Descriptie1 = Table_data1.Rows.Item(i).Item("Descriptie_custom").ToString
                        Else
                            Descriptie1 = ""
                        End If

                        Dim Block_name_string As String = ""
                        If IsDBNull(Table_data1.Rows.Item(i).Item("Block_name")) = False Then
                            Block_name_string = Table_data1.Rows.Item(i).Item("Block_name")
                        End If

                        Dim layer_name As String = ""
                        If IsDBNull(Table_data1.Rows.Item(i).Item("layer")) = False Then
                            layer_name = Table_data1.Rows.Item(i).Item("layer")
                        End If


                        Dim x, y, z As Double
                        x = East1
                        y = North1
                        If Is3D = True Then z = Elev1




                        Dim Layer1 As String
                        Dim Layer_prefix As String = ""

                        Layer_prefix = TextBox_layer_prefix.Text

                        If layer_name = "" Then
                            Layer1 = Layer_prefix & "_" & Code1
                        Else
                            Layer1 = layer_name
                        End If

                        If RadioButton_polyline_only.Checked = False And RadioButton_INSERT_Leader.Checked = False Then
                            Creaza_layer(Layer1, 3, "created by point insertor", False)
                        End If


                        Dim Text1 As New Autodesk.AutoCAD.DatabaseServices.DBText()
                        Dim Text2 As New Autodesk.AutoCAD.DatabaseServices.DBText()
                        Dim text3 As New Autodesk.AutoCAD.DatabaseServices.DBText()

                        If RadioButton_number_description_elevation.Checked = True Or
                                RadioButton_point_number_and_description.Checked = True Or
                                RadioButton_point_number_and_elevation.Checked = True Or
                                RadioButton_point_number_only.Checked = True Then

                            If Not PN1 = "" Then
                                Text1.TextString = " " & PN1
                                Text1.Position = New Autodesk.AutoCAD.Geometry.Point3d(x, y, z)
                                Text1.Height = 0.1
                                Text1.Rotation = 7 * PI / 4
                                Text1.Layer = Layer1
                                BTrecord.AppendEntity(Text1)
                                Trans1.AddNewlyCreatedDBObject(Text1, True)
                            End If


                        End If


                        If RadioButton_number_description_elevation.Checked = True Or
                                    RadioButton_point_number_and_description.Checked = True Then


                            If Not Descriptie1 = "" Then
                                Text2.TextString = Descriptie1
                                Text2.Position = New Autodesk.AutoCAD.Geometry.Point3d(x, y, z)

                                Text2.Height = 0.1
                                Text2.Rotation = 0
                                Text2.Layer = Layer1
                                BTrecord.AppendEntity(Text2)
                                Trans1.AddNewlyCreatedDBObject(Text2, True)
                            End If



                        End If

                        If RadioButton_INSERT_Leader.Checked = True Then
                            If Not Descriptie1 = "" Then
                                Creaza_Mleader_nou_fara_UCS_transform(New Autodesk.AutoCAD.Geometry.Point3d(x, y, z), Descriptie1, 2.5, 2.5, 2.5, 3, 5)
                            End If
                        End If

                        If RadioButton_number_description_elevation.Checked = True Or
                                        RadioButton_point_number_and_elevation.Checked = True Or
                                        RadioButton_Points_elevation_only.Checked = True Then

                            text3.TextString = " " & Elev1

                            If IsNumeric(TextBox_decimals.Text) = True Then
                                If CInt(TextBox_decimals.Text) >= 0 Then
                                    text3.TextString = " " & Get_String_Rounded(Elev1, CInt(TextBox_decimals.Text))
                                End If
                            End If
                            If Elev1 = 0 And IsNumeric(elev_string) = False Then
                                text3.TextString = " " & elev_string
                            End If

                            text3.Position = New Autodesk.AutoCAD.Geometry.Point3d(x, y, z)

                            text3.Height = 0.1
                            text3.Rotation = PI / 4
                            text3.Layer = Layer1
                            BTrecord.AppendEntity(text3)
                            Trans1.AddNewlyCreatedDBObject(text3, True)

                        End If

                        If RadioButton_polyline_only.Checked = False Then
                            Dim Point11 As New Autodesk.AutoCAD.DatabaseServices.DBPoint(New Autodesk.AutoCAD.Geometry.Point3d(x, y, z))
                            Point11.Layer = Layer1
                            BTrecord.AppendEntity(Point11)
                            Trans1.AddNewlyCreatedDBObject(Point11, True)
                        End If



                        Dim Index_start As New IntegerCollection

                        If CheckBox_line_code.Checked = True And Code1 > 0 And Code1 < 1000 Then
                            If Lcode1 = "1" Then
                                Start_index_line_code.Add(i)
                                Layer_line_code.Add(Layer1)
                                lista_coduri.Add(Code1)
                                Nr_de_linii = Nr_de_linii + 1
                            End If
                        End If

                        'de aici sunt bLOCURI
                        If CheckBox_insert_blocks.Checked = True And Not TextBox_block_name.Text = "" Then
                            Dim Scale1 As Double = 1
                            If IsNumeric(TextBox_block_scale.Text) = True Then
                                Scale1 = CDbl(TextBox_block_scale.Text)
                            End If
                            ' Dim Block1 As Autodesk.AutoCAD.DatabaseServices.BlockReference
                            'Block1 = InsertBlock(New Autodesk.AutoCAD.Geometry.Point3d(x, y, z), BTrecord, Block_name_string, ComboBox_Layer_for_blocks.Text, Scale1)


                            Dim Colectie_atr_name As New Specialized.StringCollection
                            Dim Colectie_atr_value As New Specialized.StringCollection



                            If IsDBNull(Table_data1.Rows(i).Item("Attribute_tag1")) = False And IsDBNull(Table_data1.Rows(i).Item("Attribute_value1")) = False Then
                                Colectie_atr_name.Add(Table_data1.Rows(i).Item("Attribute_tag1"))
                                Colectie_atr_value.Add(Table_data1.Rows(i).Item("Attribute_value1"))
                            End If

                            If IsDBNull(Table_data1.Rows(i).Item("Attribute_tag2")) = False And IsDBNull(Table_data1.Rows(i).Item("Attribute_value2")) = False Then
                                Colectie_atr_name.Add(Table_data1.Rows(i).Item("Attribute_tag2"))
                                Colectie_atr_value.Add(Table_data1.Rows(i).Item("Attribute_value2"))
                            End If

                            If Not Replace(Block_name_string, " ", "") = "" Then
                                InsertBlock_with_multiple_atributes(Block_name_string & ".dwg", Block_name_string, New Autodesk.AutoCAD.Geometry.Point3d(x, y, z), Scale1, BTrecord, ComboBox_Layer_for_blocks.Text, Colectie_atr_name, Colectie_atr_value)
                            End If
                            'asta e de la INSERT BLOCKS
                        End If


                        'asta e de la data din tabela
                    Next
                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using



                If CheckBox_line_code.Checked = True Then
                    Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                        If Button_2D_3D.Text = "2D" Then
                            For i = 0 To Nr_de_linii - 1

                                Dim Coords_x As New DoubleCollection
                                Dim Coords_y As New DoubleCollection
                                Dim Adauga_x_y As Boolean = True
                                Dim end1 As Double = CDbl(TextBox_row_end.Text)
                                For j = Start_index_line_code.Item(i) To end1
                                    Try
                                        If Table_data1.Rows.Item(j).Item("Code").ToString = lista_coduri(i).ToString Then

                                            Dim East1, North1 As Double
                                            Dim East0, North0 As Double
                                            Dim Lcode1 As Integer

                                            If IsDBNull(Table_data1.Rows.Item(i).Item("East")) = False Then
                                                If IsNumeric(Table_data1.Rows.Item(j).Item("East").ToString) = True Then
                                                    East1 = Table_data1.Rows.Item(j).Item("East")
                                                Else
                                                    East1 = 0
                                                End If
                                            Else
                                                East1 = 0
                                            End If

                                            If IsDBNull(Table_data1.Rows.Item(i).Item("North")) = False Then
                                                If IsNumeric(Table_data1.Rows.Item(j).Item("North").ToString) = True Then
                                                    North1 = Table_data1.Rows.Item(j).Item("North")
                                                Else
                                                    North1 = 0
                                                End If
                                            Else
                                                North1 = 0
                                            End If

                                            If IsDBNull(Table_data1.Rows.Item(i).Item("LCode")) = False Then
                                                If IsNumeric(Table_data1.Rows.Item(j).Item("LCode").ToString) = True Then
                                                    Lcode1 = CInt(Table_data1.Rows.Item(j).Item("LCode"))
                                                Else
                                                    Lcode1 = 0
                                                End If
                                            Else
                                                Lcode1 = 0
                                            End If

                                            If j = Start_index_line_code.Item(i) Then
                                                East0 = East1
                                                North0 = North1
                                            End If

                                            If Adauga_x_y = True Then
                                                Coords_x.Add(East1)
                                                Coords_y.Add(North1)
                                            End If


                                            If Lcode1 = 2 Then
                                                Adauga_x_y = False
                                                Exit For
                                            End If

                                            If Lcode1 = 3 Then
                                                Coords_x.Add(East0)
                                                Coords_y.Add(North0)
                                                Adauga_x_y = False
                                                Exit For
                                            End If

                                            ' asta e de la match intre coduri
                                        End If
                                    Catch ex As IndexOutOfRangeException
                                        MsgBox("Check the line code")
                                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                        Exit Sub
                                    End Try
                                    ' asta e de la fiecare linie in parte
                                Next


                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                Poly1.Layer = Layer_line_code(i)

                                For k = 0 To Coords_x.Count - 1
                                    Poly1.AddVertexAt(k, New Point2d(Coords_x.Item(k), Coords_y.Item(k)), 0, 0, 0)
                                Next

                                BTrecord.AppendEntity(Poly1)
                                Trans2.AddNewlyCreatedDBObject(Poly1, True)

                                'asta e de la nr linii
                            Next

                            ' asta e de la butonul 2d
                        End If

                        If Button_2D_3D.Text = "3D" Then
                            For i = 0 To Nr_de_linii - 1

                                Dim Coords_x As New DoubleCollection
                                Dim Coords_y As New DoubleCollection
                                Dim Coords_z As New DoubleCollection

                                Dim Adauga_x_y_z As Boolean = True
                                Dim end1 As Double = CDbl(TextBox_row_end.Text)
                                For j = Start_index_line_code.Item(i) To end1
                                    Try

                                        If Table_data1.Rows.Item(j).Item("Code").ToString = lista_coduri(i).ToString Then

                                            Dim East1, North1, Elev1 As Double
                                            Dim East0, North0, Elev0 As Double
                                            Dim Lcode1 As Integer

                                            If IsDBNull(Table_data1.Rows.Item(i).Item("East")) = False Then
                                                If IsNumeric(Table_data1.Rows.Item(j).Item("East").ToString) = True Then
                                                    East1 = Table_data1.Rows.Item(j).Item("East")
                                                Else
                                                    East1 = 0
                                                End If
                                            Else
                                                East1 = 0
                                            End If

                                            If IsDBNull(Table_data1.Rows.Item(i).Item("North")) = False Then
                                                If IsNumeric(Table_data1.Rows.Item(j).Item("North").ToString) = True Then
                                                    North1 = Table_data1.Rows.Item(j).Item("North")
                                                Else
                                                    North1 = 0
                                                End If
                                            Else
                                                North1 = 0
                                            End If

                                            If IsDBNull(Table_data1.Rows.Item(i).Item("Elevation")) = False Then
                                                If IsNumeric(Table_data1.Rows.Item(j).Item("Elevation").ToString) = True Then
                                                    Elev1 = Table_data1.Rows.Item(j).Item("Elevation")
                                                Else
                                                    Elev1 = 0
                                                End If
                                            Else
                                                Elev1 = 0
                                            End If

                                            If IsDBNull(Table_data1.Rows.Item(i).Item("LCode")) = False Then
                                                If IsNumeric(Table_data1.Rows.Item(j).Item("LCode").ToString) = True Then
                                                    Lcode1 = CInt(Table_data1.Rows.Item(j).Item("LCode"))
                                                Else
                                                    Lcode1 = 0
                                                End If
                                            Else
                                                Lcode1 = 0
                                            End If

                                            If j = Start_index_line_code.Item(i) Then
                                                East0 = East1
                                                North0 = North1
                                                Elev0 = Elev1
                                            End If

                                            If Adauga_x_y_z = True Then
                                                Coords_x.Add(East1)
                                                Coords_y.Add(North1)
                                                Coords_z.Add(Elev1)
                                            End If


                                            If Lcode1 = 2 Then
                                                Adauga_x_y_z = False
                                                Exit For
                                            End If

                                            If Lcode1 = 3 Then
                                                Coords_x.Add(East0)
                                                Coords_y.Add(North0)
                                                Coords_z.Add(Elev0)
                                                Adauga_x_y_z = False
                                                Exit For
                                            End If

                                            ' asta e de la match intre coduri
                                        End If
                                    Catch ex As IndexOutOfRangeException
                                        MsgBox("Check the line code")
                                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                        Exit Sub
                                    End Try
                                    ' asta e de la fiecare linie in parte
                                Next


                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline3d
                                Poly1.SetDatabaseDefaults()
                                Poly1.Layer = Layer_line_code(i)

                                BTrecord.AppendEntity(Poly1)
                                Trans2.AddNewlyCreatedDBObject(Poly1, True)

                                Dim Colectie_puncte_3d As New Point3dCollection

                                For k = 0 To Coords_x.Count - 1
                                    Colectie_puncte_3d.Add(New Point3d(Coords_x.Item(k), Coords_y.Item(k), Coords_z.Item(k)))
                                Next

                                For Each punct As Point3d In Colectie_puncte_3d
                                    Dim Vertex_poly_3d As New PolylineVertex3d(punct)
                                    Poly1.AppendVertex(Vertex_poly_3d)
                                    Trans2.AddNewlyCreatedDBObject(Vertex_poly_3d, True)
                                Next


                                'asta e de la nr linii
                            Next

                            ' asta e de la butonul 3d
                        End If

                        Trans2.Commit()

                        'ASTA E DE LA LINE CODES
                    End Using

                    ' asta e de la check line codes . checked = true
                End If


                If RadioButton_polyline_only.Checked = True Then
                    Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                        If Button_2D_3D.Text = "2D" Then

                            If Table_data1.Rows.Count > 1 Then
                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                                Poly1.Layer = ComboBox_poly_layer.Text

                                For i = 0 To Table_data1.Rows.Count - 1
                                    If IsDBNull(Table_data1.Rows.Item(i).Item("East")) = False And IsDBNull(Table_data1.Rows.Item(i).Item("North")) = False Then
                                        Dim X As Double = Table_data1.Rows.Item(i).Item("East")
                                        Dim y As Double = Table_data1.Rows.Item(i).Item("North")
                                        Poly1.AddVertexAt(i, New Point2d(X, y), 0, 0, 0)
                                    End If
                                Next

                                BTrecord.AppendEntity(Poly1)
                                Trans2.AddNewlyCreatedDBObject(Poly1, True)
                            End If



                            ' asta e de la butonul 2d
                        End If

                        If Button_2D_3D.Text = "3D" Then


                            If Table_data1.Rows.Count > 1 Then
                                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline3d
                                Poly1.SetDatabaseDefaults()
                                Poly1.Layer = ComboBox_poly_layer.Text

                                BTrecord.AppendEntity(Poly1)
                                Trans2.AddNewlyCreatedDBObject(Poly1, True)

                                Dim Colectie_puncte_3d As New Point3dCollection

                                For i = 0 To Table_data1.Rows.Count - 1
                                    If IsDBNull(Table_data1.Rows.Item(i).Item("East")) = False And IsDBNull(Table_data1.Rows.Item(i).Item("North")) = False And IsDBNull(Table_data1.Rows.Item(i).Item("Elevation")) = False Then
                                        Dim X As Double = Table_data1.Rows.Item(i).Item("East")
                                        Dim y As Double = Table_data1.Rows.Item(i).Item("North")
                                        Dim Z As Double = Table_data1.Rows.Item(i).Item("Elevation")
                                        Colectie_puncte_3d.Add(New Point3d(X, y, Z))
                                    End If

                                Next

                                If Colectie_puncte_3d.Count > 0 Then
                                    For Each punct As Point3d In Colectie_puncte_3d
                                        Dim Vertex_poly_3d As New PolylineVertex3d(punct)
                                        Poly1.AppendVertex(Vertex_poly_3d)
                                        Trans2.AddNewlyCreatedDBObject(Vertex_poly_3d, True)
                                    Next
                                End If
                            End If



                            ' asta e de la butonul 3d
                        End If

                        Trans2.Commit()

                        'ASTA E DE LA LINE CODES
                    End Using

                    ' asta e de la check line codes . checked = true
                End If


                afiseaza_butoanele_pentru_forms(Me, Colectie1)


                TextBox_message.Text = "Points inserted"
                ThisDrawing.Editor.WriteMessage(vbCrLf & "Command:")

                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                ' asta e de la lock
            End Using

            Exit Sub
123:

            afiseaza_butoanele_pentru_forms(Me, Colectie1)

        Catch ex As Exception

            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub Panel_COLUMNS_Click(sender As Object, e As EventArgs) Handles Panel_BLOCKS.Click
        Incarca_existing_layers_to_combobox(ComboBox_Layer_for_blocks)
        ComboBox_Layer_for_blocks.SelectedIndex = 0
        If ComboBox_Layer_for_blocks.Items.Contains("TEXT") = True Then
            ComboBox_Layer_for_blocks.SelectedIndex = ComboBox_Layer_for_blocks.Items.IndexOf("TEXT")
        End If
        If ComboBox_Layer_for_blocks.Items.Contains("Text") = True Then
            ComboBox_Layer_for_blocks.SelectedIndex = ComboBox_Layer_for_blocks.Items.IndexOf("Text")
        End If
        If ComboBox_Layer_for_blocks.Items.Contains("text") = True Then
            ComboBox_Layer_for_blocks.SelectedIndex = ComboBox_Layer_for_blocks.Items.IndexOf("text")
        End If
    End Sub

    Private Sub Panel_points_Click(sender As Object, e As EventArgs) Handles Panel_points.Click
        Incarca_existing_layers_to_combobox(ComboBox_poly_layer)
        ComboBox_poly_layer.SelectedIndex = 0
    End Sub
    Private Sub TextBox_Point_name_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Point_name.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_East
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_East_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_East.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_NORTH
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_NORTH_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_NORTH.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_elevation
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_elevation_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_elevation.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_description
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_description_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_description.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_extra1
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_code_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_extra1.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_extra2
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_Line_code_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_extra2.KeyDown, TextBox_ln.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_row_start
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_row_start_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_row_start.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With TextBox_row_end
                .SelectAll()
                .Focus()
            End With

        End If
    End Sub

    Private Sub TextBox_row_end_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox_row_end.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            With CheckBox_line_code
                If .Checked = True Then
                    .Checked = False
                Else
                    .Checked = True
                End If


            End With

        End If
    End Sub




    Private Sub Button_w2xl_Click(sender As Object, e As EventArgs)
        Try


            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database = ThisDrawing.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim dt1 As New System.Data.DataTable
            dt1.Columns.Add("x", GetType(Double))
            dt1.Columns.Add("y", GetType(Double))
            dt1.Columns.Add("z", GetType(Double))
            dt1.Columns.Add("text", GetType(String))

            Using Lock1 As DocumentLock = ThisDrawing.LockDocument

                Dim Lista_layere_data As New Specialized.StringCollection


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Dim Descr1 As String = "created by point insertor"

                    Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                    LayerTable1 = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim Lista_layere_de_scanat As New Specialized.StringCollection

                    For Each Obj_Id1 As ObjectId In LayerTable1
                        Dim Layer1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                        Layer1 = Trans1.GetObject(Obj_Id1, OpenMode.ForRead)
                        If Layer1.Description.ToUpper = Descr1.ToUpper Then
                            Lista_layere_de_scanat.Add(Layer1.Name)
                        End If
                    Next



                    For Each od1 As ObjectId In BTrecord


                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(od1, OpenMode.ForRead), Entity)
                        If IsNothing(Ent1) = False Then
                            If Lista_layere_de_scanat.Contains(Ent1.Layer) Then
                                If TypeOf Ent1 Is DBText Then
                                    Dim txt1 As DBText = Ent1
                                    dt1.Rows.Add()
                                    dt1.Rows(dt1.Rows.Count - 1).Item("x") = txt1.Position.X
                                    dt1.Rows(dt1.Rows.Count - 1).Item("y") = txt1.Position.Y
                                    dt1.Rows(dt1.Rows.Count - 1).Item("z") = txt1.Position.Z
                                    dt1.Rows(dt1.Rows.Count - 1).Item("text") = txt1.TextString

                                End If

                            End If
                        End If



                    Next

                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using
                ' asta e de la lock
            End Using

            Transfer_datatable_to_new_excel_spreadsheet(dt1)

            TextBox_message.Text = "done"

            ' asta e de la mesage box


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub


End Class