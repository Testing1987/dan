Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class FIND_REPLACE_Form

    Private Sub Global_change_Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ComboBox_color.Items.Add("BYLAYER")
        ComboBox_color.Items.Add("255,0,0")
        ComboBox_color.SelectedIndex = 0
    End Sub

    Private Sub Button_Load_Info_from_excel_Click(sender As Object, e As EventArgs) Handles Button_Load_Info_from_excel.Click
        Try
            If TextBox_FIND_col_xl.Text = "" Then
                MsgBox("Please specify the EXCEL COLUMN!")
                Exit Sub
            End If
            If TextBox_REPLACE_COL_XL.Text = "" Then
                MsgBox("Please specify the EXCEL COLUMN!")
                Exit Sub
            End If

            If Val(TextBox_ROW_START.Text) < 1 Or IsNumeric(TextBox_ROW_START.Text) = False Then
                With TextBox_ROW_START
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If

            If Val(TextBox_ROW_END.Text) < 1 Or IsNumeric(TextBox_ROW_END.Text) = False Then
                With TextBox_ROW_END
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")

                Exit Sub
            End If
            If Val(TextBox_ROW_END.Text) < Val(TextBox_ROW_START.Text) Then
                With TextBox_ROW_END
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row smaller than start row")

                Exit Sub
            End If
            Ascunde_butoanele()
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_ROW_START.Text)
            Dim end1 As Integer = CInt(TextBox_ROW_END.Text)
            ListBox_FIND.Items.Clear()
            ListBox_REPLACE.Items.Clear()
            For i = start1 To end1
                Dim Find1 As String = W1.Range(TextBox_FIND_col_xl.Text.ToUpper & i).Text
                Dim Replace1 As String = W1.Range(TextBox_REPLACE_COL_XL.Text.ToUpper & i).Text
                If Not Replace(Find1, " ", "") = "" And Not Replace(Replace1, " ", "") = "" Then
                    ListBox_FIND.Items.Add(Find1)
                    ListBox_REPLACE.Items.Add(Replace1)
                End If

            Next
            Show_butoanele()
        Catch ex As Exception
            Show_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub Ascunde_butoanele()
        For i = 0 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is Windows.Forms.Button Then
                Me.Controls(i).Visible = False
            End If

        Next
    End Sub
    Public Sub Show_butoanele()

        For i = 0 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is Windows.Forms.Button Then
                Me.Controls(i).Visible = True
            End If

        Next
    End Sub

    Private Sub Button_EXECUTE_BACKGROUND_Click(sender As System.Object, e As System.EventArgs) Handles Button_EXECUTE.Click
        Try

            If TextBox_source_folder.Text = "" Then
                MsgBox("Please specify the source folder!")
                Exit Sub
            End If
            If TextBox_destination_folder.Text = "" Then
                MsgBox("Please specify the destination folder!")
                Exit Sub
            End If
            Dim Folder_sursa As String = TextBox_source_folder.Text
            Dim Folder_destinatie As String = TextBox_destination_folder.Text

            If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Folder_sursa) = False Then
                MsgBox("The source folder doesn't exists!")
                Exit Sub
            End If

            If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Folder_destinatie) = False Then
                System.IO.Directory.CreateDirectory(Folder_destinatie)
            End If

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()

            Ascunde_butoanele()

            Dim Colectie_nume_fisiere As New Specialized.StringCollection

            For Each File1 As String In My.Computer.FileSystem.GetFiles(Folder_sursa, FileIO.SearchOption.SearchAllSubDirectories, "*.dwg")
                Colectie_nume_fisiere.Add(File1)
            Next
            'If System.IO.File.Exists(Director_processed & "\" & Table_data.Rows.Item(i).Item("Col_New_file_name")) = False Then
            For j = 0 To Colectie_nume_fisiere.Count - 1
                Dim Fisier As String = Colectie_nume_fisiere(j)
                W1.Cells(j + 1, 1).VALUE = Fisier

                Dim Database1 As New Database(False, True)
                Database1.ReadDwgFile(Fisier, IO.FileShare.ReadWrite, False, Nothing)
                Dim Este_gasit As Boolean = False


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction

                    Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                    For Each Id1 As ObjectId In BlockTable1
                        Dim Blocktablerecord1 As BlockTableRecord = Id1.GetObject(OpenMode.ForRead)
                        If Blocktablerecord1.IsLayout = True Then
                            If Not Blocktablerecord1.Name.ToLower = "*model_space" Then
                                Blocktablerecord1.UpgradeOpen()
                                For Each Id2 As ObjectId In Blocktablerecord1
                                    Dim Ent1 As Entity = Id2.GetObject(OpenMode.ForRead)

                                    If TypeOf Ent1 Is DBText Then
                                        Try
                                            Dim Text1 As DBText = Ent1
                                            Dim String1 As String = Text1.TextString
                                            For i = 0 To ListBox_FIND.Items.Count - 1
                                                Dim Find1 As String = ListBox_FIND.Items(i)
                                                Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                    If InStr(String1, Find1) > 0 Then
                                                        Text1.UpgradeOpen()
                                                        Text1.TextString = Replace(String1, Find1, Replace1)
                                                        Este_gasit = True
                                                    End If
                                                End If

                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                    String1 = Text1.TextString
                                                    Dim Word1 As String = " " & Find1 & " "
                                                    If InStr(String1, Word1) > 0 Then
                                                        Text1.UpgradeOpen()
                                                        Text1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                                        Este_gasit = True
                                                    End If
                                                    String1 = Text1.TextString
                                                    Word1 = " " & Find1
                                                    Dim Lungime_word1 As Integer = Len(Word1)

                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                        Text1.UpgradeOpen()
                                                        Text1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                        Este_gasit = True
                                                    End If
                                                    String1 = Text1.TextString
                                                    Word1 = Find1 & " "
                                                    Lungime_word1 = Len(Word1)
                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                        Text1.UpgradeOpen()
                                                        Text1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                        Este_gasit = True
                                                    End If

                                                    If String1 = Find1 Then
                                                        Text1.UpgradeOpen()
                                                        Text1.TextString = Replace1
                                                        Este_gasit = True
                                                    End If
                                                End If
                                            Next
                                        Catch ex As System.Exception

                                        End Try

                                    End If

                                    If TypeOf Ent1 Is MText Then
                                        Try
                                            Dim MText1 As MText = Ent1

                                            For i = 0 To ListBox_FIND.Items.Count - 1
                                                Dim Find1 As String = ListBox_FIND.Items(i)
                                                Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                                Dim String1 As String = MText1.Contents

                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                    If InStr(String1, Find1) > 0 Then
                                                        Dim String2 As String = Replace(String1, Find1, Replace1)
                                                        MText1.UpgradeOpen()
                                                        MText1.Contents = String2
                                                        Este_gasit = True
                                                    End If
                                                End If


                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                    String1 = MText1.Contents
                                                    Dim Word1 As String = " " & Find1 & " "
                                                    If InStr(String1, Word1) > 0 Then
                                                        MText1.UpgradeOpen()
                                                        MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                        Este_gasit = True
                                                    End If
                                                    String1 = MText1.Contents
                                                    Word1 = " " & Find1
                                                    Dim Lungime_word1 As Integer = Len(Word1)

                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                        MText1.UpgradeOpen()
                                                        MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                        Este_gasit = True
                                                    End If
                                                    String1 = MText1.Contents
                                                    Word1 = Find1 & " "
                                                    Lungime_word1 = Len(Word1)
                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                        MText1.UpgradeOpen()
                                                        MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                        Este_gasit = True
                                                    End If

                                                    If String1 = Find1 Then
                                                        MText1.UpgradeOpen()
                                                        MText1.Contents = Replace1
                                                        Este_gasit = True
                                                    End If
                                                End If

                                            Next
                                        Catch ex As System.Exception
                                        End Try

                                    End If

                                    If TypeOf Ent1 Is MLeader Then
                                        Try
                                            Dim Mleader1 As MLeader = Ent1
                                            For i = 0 To ListBox_FIND.Items.Count - 1
                                                Dim MText1 As MText = Mleader1.MText
                                                Dim String1 As String = MText1.Contents
                                                Dim Find1 As String = ListBox_FIND.Items(i)
                                                Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                    If InStr(String1, Find1) > 0 Then
                                                        Mleader1.UpgradeOpen()
                                                        MText1.Contents = Replace(String1, Find1, Replace1)
                                                        Mleader1.MText = MText1
                                                        Este_gasit = True
                                                    End If
                                                End If


                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                    String1 = MText1.Contents
                                                    Dim Word1 As String = " " & Find1 & " "
                                                    If InStr(String1, Word1) > 0 Then
                                                        Mleader1.UpgradeOpen()
                                                        MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                        Mleader1.MText = MText1
                                                        Este_gasit = True
                                                    End If

                                                    String1 = MText1.Contents
                                                    Word1 = " " & Find1
                                                    Dim Lungime_word1 As Integer = Len(Word1)

                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                        Mleader1.UpgradeOpen()
                                                        MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                        Mleader1.MText = MText1
                                                        Este_gasit = True
                                                    End If
                                                    String1 = MText1.Contents
                                                    Word1 = Find1 & " "
                                                    Lungime_word1 = Len(Word1)
                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                        Mleader1.UpgradeOpen()
                                                        MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                        Mleader1.MText = MText1
                                                        Este_gasit = True
                                                    End If

                                                    If String1 = Find1 Then
                                                        Mleader1.UpgradeOpen()
                                                        MText1.Contents = Replace1
                                                        Mleader1.MText = MText1
                                                        Este_gasit = True
                                                    End If
                                                End If

                                            Next
                                        Catch ex As System.Exception
                                        End Try

                                    End If

                                    If TypeOf Ent1 Is BlockReference Then
                                        Try
                                            Dim Block1 As BlockReference = Ent1
                                            For i = 0 To ListBox_FIND.Items.Count - 1
                                                Dim Find1 As String = ListBox_FIND.Items(i)
                                                Dim Replace1 As String = ListBox_REPLACE.Items(i)

                                                If Block1.AttributeCollection.Count > 0 Then
                                                    For Each Atid As ObjectId In Block1.AttributeCollection
                                                        Dim Atr1 As AttributeReference = Atid.GetObject(OpenMode.ForRead)
                                                        If Atr1.IsMTextAttribute = True Then
                                                            Dim String1 As String = Atr1.MTextAttribute.Contents
                                                            If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                                If InStr(String1, Find1) > 0 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                    String1 = Atr1.MTextAttribute.Contents
                                                                    Mtext1.Contents = Replace(String1, Find1, Replace1)
                                                                    Atr1.MTextAttribute = Mtext1
                                                                    Este_gasit = True
                                                                End If
                                                            End If

                                                            If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                                String1 = Atr1.MTextAttribute.Contents
                                                                Dim Word1 As String = " " & Find1 & " "
                                                                If InStr(String1, Word1) > 0 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                    Mtext1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                                    Atr1.MTextAttribute = Mtext1
                                                                    Este_gasit = True
                                                                End If

                                                                String1 = Atr1.MTextAttribute.Contents
                                                                Word1 = " " & Find1
                                                                Dim Lungime_word1 As Integer = Len(Word1)

                                                                If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                    Mtext1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                                    Atr1.MTextAttribute = Mtext1
                                                                    Este_gasit = True
                                                                End If

                                                                String1 = Atr1.MTextAttribute.Contents
                                                                Word1 = Find1 & " "
                                                                Lungime_word1 = Len(Word1)
                                                                If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                    Mtext1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                                    Atr1.MTextAttribute = Mtext1
                                                                    Este_gasit = True
                                                                End If

                                                                If String1 = Find1 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                    Mtext1.Contents = Replace1
                                                                    Atr1.MTextAttribute = Mtext1
                                                                    Este_gasit = True
                                                                End If
                                                            End If
                                                        Else
                                                            Dim String1 As String = Atr1.TextString
                                                            If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                                If InStr(String1, Find1) > 0 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Atr1.TextString = Replace(String1, Find1, Replace1)
                                                                    Este_gasit = True
                                                                End If
                                                            End If

                                                            If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                                String1 = Atr1.TextString
                                                                Dim Word1 As String = " " & Find1 & " "
                                                                If InStr(String1, Word1) > 0 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Atr1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                                                    Este_gasit = True
                                                                End If

                                                                String1 = Atr1.TextString
                                                                Word1 = " " & Find1
                                                                Dim Lungime_word1 As Integer = Len(Word1)

                                                                If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Atr1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                                    Este_gasit = True
                                                                End If

                                                                String1 = Atr1.TextString
                                                                Word1 = Find1 & " "
                                                                Lungime_word1 = Len(Word1)
                                                                If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Atr1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                                    Este_gasit = True
                                                                End If

                                                                If String1 = Find1 Then
                                                                    Atr1.UpgradeOpen()
                                                                    Atr1.TextString = Replace1
                                                                    Este_gasit = True
                                                                End If
                                                            End If


                                                        End If
                                                    Next

                                                End If

                                            Next
                                        Catch ex As System.Exception
                                        End Try

                                    End If
                                Next
                            End If

                            If CheckBox_search_modelspace.Checked = True Then
                                If Blocktablerecord1.Name.ToLower = "*model_space" Then
                                    Blocktablerecord1.UpgradeOpen()
                                    For Each Id2 As ObjectId In Blocktablerecord1
                                        Dim Ent1 As Entity = Id2.GetObject(OpenMode.ForRead)

                                        If TypeOf Ent1 Is DBText Then
                                            Try
                                                Dim Text1 As DBText = Ent1
                                                Dim String1 As String = Text1.TextString
                                                For i = 0 To ListBox_FIND.Items.Count - 1
                                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                        If InStr(String1, Find1) > 0 Then
                                                            Text1.UpgradeOpen()
                                                            Text1.TextString = Replace(String1, Find1, Replace1)
                                                            Este_gasit = True
                                                        End If
                                                    End If

                                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                        String1 = Text1.TextString
                                                        Dim Word1 As String = " " & Find1 & " "
                                                        If InStr(String1, Word1) > 0 Then
                                                            Text1.UpgradeOpen()
                                                            Text1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                                            Este_gasit = True
                                                        End If
                                                        String1 = Text1.TextString
                                                        Word1 = " " & Find1
                                                        Dim Lungime_word1 As Integer = Len(Word1)

                                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                            Text1.UpgradeOpen()
                                                            Text1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                            Este_gasit = True
                                                        End If
                                                        String1 = Text1.TextString
                                                        Word1 = Find1 & " "
                                                        Lungime_word1 = Len(Word1)
                                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                            Text1.UpgradeOpen()
                                                            Text1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                            Este_gasit = True
                                                        End If

                                                        If String1 = Find1 Then
                                                            Text1.UpgradeOpen()
                                                            Text1.TextString = Replace1
                                                            Este_gasit = True
                                                        End If
                                                    End If
                                                Next
                                            Catch ex As System.Exception

                                            End Try

                                        End If

                                        If TypeOf Ent1 Is MText Then
                                            Try
                                                Dim MText1 As MText = Ent1

                                                For i = 0 To ListBox_FIND.Items.Count - 1
                                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                                    Dim String1 As String = MText1.Contents

                                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                        If InStr(String1, Find1) > 0 Then
                                                            Dim String2 As String = Replace(String1, Find1, Replace1)
                                                            MText1.UpgradeOpen()
                                                            MText1.Contents = String2
                                                            Este_gasit = True
                                                        End If
                                                    End If


                                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                        String1 = MText1.Contents
                                                        Dim Word1 As String = " " & Find1 & " "
                                                        If InStr(String1, Word1) > 0 Then
                                                            MText1.UpgradeOpen()
                                                            MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                            Este_gasit = True
                                                        End If
                                                        String1 = MText1.Contents
                                                        Word1 = " " & Find1
                                                        Dim Lungime_word1 As Integer = Len(Word1)

                                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                            MText1.UpgradeOpen()
                                                            MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                            Este_gasit = True
                                                        End If
                                                        String1 = MText1.Contents
                                                        Word1 = Find1 & " "
                                                        Lungime_word1 = Len(Word1)
                                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                            MText1.UpgradeOpen()
                                                            MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                            Este_gasit = True
                                                        End If

                                                        If String1 = Find1 Then
                                                            MText1.UpgradeOpen()
                                                            MText1.Contents = Replace1
                                                            Este_gasit = True
                                                        End If
                                                    End If

                                                Next
                                            Catch ex As System.Exception
                                            End Try

                                        End If

                                        If TypeOf Ent1 Is MLeader Then
                                            Try
                                                Dim Mleader1 As MLeader = Ent1
                                                For i = 0 To ListBox_FIND.Items.Count - 1
                                                    Dim MText1 As MText = Mleader1.MText
                                                    Dim String1 As String = MText1.Contents
                                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                        If InStr(String1, Find1) > 0 Then
                                                            Mleader1.UpgradeOpen()
                                                            MText1.Contents = Replace(String1, Find1, Replace1)
                                                            Mleader1.MText = MText1
                                                            Este_gasit = True
                                                        End If
                                                    End If


                                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                        String1 = MText1.Contents
                                                        Dim Word1 As String = " " & Find1 & " "
                                                        If InStr(String1, Word1) > 0 Then
                                                            Mleader1.UpgradeOpen()
                                                            MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                            Mleader1.MText = MText1
                                                            Este_gasit = True
                                                        End If

                                                        String1 = MText1.Contents
                                                        Word1 = " " & Find1
                                                        Dim Lungime_word1 As Integer = Len(Word1)

                                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                            Mleader1.UpgradeOpen()
                                                            MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                            Mleader1.MText = MText1
                                                            Este_gasit = True
                                                        End If
                                                        String1 = MText1.Contents
                                                        Word1 = Find1 & " "
                                                        Lungime_word1 = Len(Word1)
                                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                            Mleader1.UpgradeOpen()
                                                            MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                            Mleader1.MText = MText1
                                                            Este_gasit = True
                                                        End If

                                                        If String1 = Find1 Then
                                                            Mleader1.UpgradeOpen()
                                                            MText1.Contents = Replace1
                                                            Mleader1.MText = MText1
                                                            Este_gasit = True
                                                        End If
                                                    End If

                                                Next
                                            Catch ex As System.Exception
                                            End Try

                                        End If

                                        If TypeOf Ent1 Is BlockReference Then
                                            Try
                                                Dim Block1 As BlockReference = Ent1
                                                For i = 0 To ListBox_FIND.Items.Count - 1
                                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)

                                                    If Block1.AttributeCollection.Count > 0 Then
                                                        For Each Atid As ObjectId In Block1.AttributeCollection
                                                            Dim Atr1 As AttributeReference = Atid.GetObject(OpenMode.ForRead)
                                                            If Atr1.IsMTextAttribute = True Then
                                                                Dim String1 As String = Atr1.MTextAttribute.Contents
                                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                                    If InStr(String1, Find1) > 0 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                        String1 = Atr1.MTextAttribute.Contents
                                                                        Mtext1.Contents = Replace(String1, Find1, Replace1)
                                                                        Atr1.MTextAttribute = Mtext1
                                                                        Este_gasit = True
                                                                    End If
                                                                End If

                                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                                    String1 = Atr1.MTextAttribute.Contents
                                                                    Dim Word1 As String = " " & Find1 & " "
                                                                    If InStr(String1, Word1) > 0 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                        Mtext1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                                        Atr1.MTextAttribute = Mtext1
                                                                        Este_gasit = True
                                                                    End If

                                                                    String1 = Atr1.MTextAttribute.Contents
                                                                    Word1 = " " & Find1
                                                                    Dim Lungime_word1 As Integer = Len(Word1)

                                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                        Mtext1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                                        Atr1.MTextAttribute = Mtext1
                                                                        Este_gasit = True
                                                                    End If

                                                                    String1 = Atr1.MTextAttribute.Contents
                                                                    Word1 = Find1 & " "
                                                                    Lungime_word1 = Len(Word1)
                                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                        Mtext1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                                        Atr1.MTextAttribute = Mtext1
                                                                        Este_gasit = True
                                                                    End If

                                                                    If String1 = Find1 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                                        Mtext1.Contents = Replace1
                                                                        Atr1.MTextAttribute = Mtext1
                                                                        Este_gasit = True
                                                                    End If
                                                                End If
                                                            Else
                                                                Dim String1 As String = Atr1.TextString
                                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                                    If InStr(String1, Find1) > 0 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Atr1.TextString = Replace(String1, Find1, Replace1)
                                                                        Este_gasit = True
                                                                    End If
                                                                End If

                                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                                    String1 = Atr1.TextString
                                                                    Dim Word1 As String = " " & Find1 & " "
                                                                    If InStr(String1, Word1) > 0 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Atr1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                                                        Este_gasit = True
                                                                    End If

                                                                    String1 = Atr1.TextString
                                                                    Word1 = " " & Find1
                                                                    Dim Lungime_word1 As Integer = Len(Word1)

                                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Atr1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                                        Este_gasit = True
                                                                    End If

                                                                    String1 = Atr1.TextString
                                                                    Word1 = Find1 & " "
                                                                    Lungime_word1 = Len(Word1)
                                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Atr1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                                        Este_gasit = True
                                                                    End If

                                                                    If String1 = Find1 Then
                                                                        Atr1.UpgradeOpen()
                                                                        Atr1.TextString = Replace1
                                                                        Este_gasit = True
                                                                    End If
                                                                End If


                                                            End If
                                                        Next

                                                    End If

                                                Next
                                            Catch ex As System.Exception
                                            End Try

                                        End If


                                    Next
                                End If
                            End If
                        End If

                    Next





                    If Este_gasit = True Then

                        Dim Nume_nou As String = Folder_destinatie & "\" & System.IO.Path.GetFileNameWithoutExtension(Database1.Filename) & ".dwg"
                        Database1.SaveAs(Nume_nou, Database1.OriginalFileVersion)
                    Else


                        If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Folder_destinatie & "\NOT FOUND") = False Then
                            System.IO.Directory.CreateDirectory(Folder_destinatie & "\NOT FOUND")
                        End If
                        Dim Nume_nou_NEGASIT As String = Folder_destinatie & "\NOT FOUND" & "\" & System.IO.Path.GetFileNameWithoutExtension(Database1.Filename) & ".dwg"
                        Database1.SaveAs(Nume_nou_NEGASIT, Database1.OriginalFileVersion)
                    End If

                End Using ' asta e de la tranzactie
            Next ' asta e de la For i = 0 To Colectie_nume_fisiere.Count - 1

            Show_butoanele()
            MsgBox("Finished")


            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            Show_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_current_dwg_Click(sender As Object, e As EventArgs) Handles Button_current_dwg.Click
        Try
            Ascunde_butoanele()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Editor1 = ThisDrawing.Editor
            Dim Nr_de_obiecte As Integer
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                    Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                    Dim Blocktablerecord1 As BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    For Each Id2 As ObjectId In Blocktablerecord1
                        Dim Ent1 As Entity = Id2.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Try
                                Dim Text1 As DBText = Ent1
                                Dim String1 As String = Text1.TextString
                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                        If InStr(String1, Find1) > 0 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace(String1, Find1, Replace1)
                                            If ComboBox_color.Text = "255,0,0" Then Text1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)


                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If
                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                        String1 = Text1.TextString
                                        Dim Word1 As String = " " & Find1 & " "
                                        If InStr(String1, Word1) > 0 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = Text1.TextString
                                        Word1 = " " & Find1
                                        Dim Lungime_word1 As Integer = Len(Word1)

                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = Text1.TextString
                                        Word1 = Find1 & " "
                                        Lungime_word1 = Len(Word1)
                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        If String1 = Find1 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If
                                Next
                            Catch ex As System.Exception

                            End Try
                        End If
                        If TypeOf Ent1 Is MText Then
                            Try
                                Dim MText1 As MText = Ent1

                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                    Dim String1 As String = MText1.Contents

                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                        If InStr(String1, Find1) > 0 Then
                                            Dim String2 As String = Replace(String1, Find1, Replace1)
                                            MText1.UpgradeOpen()
                                            MText1.Contents = String2
                                            If ComboBox_color.Text = "255,0,0" Then MText1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If


                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                        String1 = MText1.Contents
                                        Dim Word1 As String = " " & Find1 & " "
                                        If InStr(String1, Word1) > 0 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = MText1.Contents
                                        Word1 = " " & Find1
                                        Dim Lungime_word1 As Integer = Len(Word1)

                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = MText1.Contents
                                        Word1 = Find1 & " "
                                        Lungime_word1 = Len(Word1)
                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        If String1 = Find1 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If

                                Next
                            Catch ex As System.Exception
                            End Try

                        End If

                        If TypeOf Ent1 Is MLeader Then
                            Try
                                Dim Mleader1 As MLeader = Ent1
                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim MText1 As MText = Mleader1.MText
                                    Dim String1 As String = MText1.Contents
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                        If InStr(String1, Find1) > 0 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace(String1, Find1, Replace1)
                                            If ComboBox_color.Text = "255,0,0" Then MText1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1

                                        End If
                                    End If


                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                        String1 = MText1.Contents
                                        Dim Word1 As String = " " & Find1 & " "
                                        If InStr(String1, Word1) > 0 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        String1 = MText1.Contents
                                        Word1 = " " & Find1
                                        Dim Lungime_word1 As Integer = Len(Word1)

                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = MText1.Contents
                                        Word1 = Find1 & " "
                                        Lungime_word1 = Len(Word1)
                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        If String1 = Find1 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace1
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If

                                Next
                            Catch ex As System.Exception
                            End Try

                        End If

                        If TypeOf Ent1 Is BlockReference Then
                            Try
                                Dim Block1 As BlockReference = Ent1
                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)

                                    If Block1.AttributeCollection.Count > 0 Then
                                        For Each Atid As ObjectId In Block1.AttributeCollection
                                            Dim Atr1 As AttributeReference = Atid.GetObject(OpenMode.ForRead)
                                            If Atr1.IsMTextAttribute = True Then
                                                Dim String1 As String = Atr1.MTextAttribute.Contents
                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                    If InStr(String1, Find1) > 0 Then
                                                        Atr1.UpgradeOpen()

                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        String1 = Atr1.MTextAttribute.Contents
                                                        Mtext1.Contents = Replace(String1, Find1, Replace1)
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If

                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                    String1 = Atr1.MTextAttribute.Contents
                                                    Dim Word1 As String = " " & Find1 & " "
                                                    If InStr(String1, Word1) > 0 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    String1 = Atr1.MTextAttribute.Contents
                                                    Word1 = " " & Find1
                                                    Dim Lungime_word1 As Integer = Len(Word1)

                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    String1 = Atr1.MTextAttribute.Contents
                                                    Word1 = Find1 & " "
                                                    Lungime_word1 = Len(Word1)
                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    If String1 = Find1 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Replace1
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If
                                            Else
                                                Dim String1 As String = Atr1.TextString
                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                    If InStr(String1, Find1) > 0 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace(String1, Find1, Replace1)
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If

                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                    String1 = Atr1.TextString
                                                    Dim Word1 As String = " " & Find1 & " "
                                                    If InStr(String1, Word1) > 0 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    String1 = Atr1.TextString
                                                    Word1 = " " & Find1
                                                    Dim Lungime_word1 As Integer = Len(Word1)
                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                    String1 = Atr1.TextString
                                                    Word1 = Find1 & " "
                                                    Lungime_word1 = Len(Word1)
                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                    If String1 = Find1 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If
                                Next
                            Catch ex As System.Exception
                            End Try
                        End If
                    Next


                    Trans1.Commit()
                End Using ' asta e de la tranzactie
            End Using
            Show_butoanele()
            MsgBox("Finished" & vbCrLf & Nr_de_obiecte & " items replaced")


            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            Show_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_FIND_REPLACE_SELECTION_Click(sender As Object, e As EventArgs) Handles Button_FIND_REPLACE_SELECTION.Click
        Try
            Ascunde_butoanele()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Editor1 = ThisDrawing.Editor
            Dim Nr_de_obiecte As Integer
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                    Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                    Dim Blocktablerecord1 As BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Show_butoanele()
                        Exit Sub
                    End If
                    If Rezultat1.Value.Count < 1 Then
                        Show_butoanele()
                        Exit Sub
                    End If

                    For indx = 0 To Rezultat1.Value.Count - 1
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(indx)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Try
                                Dim Text1 As DBText = Ent1
                                Dim String1 As String = Text1.TextString
                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                        If InStr(String1, Find1) > 0 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace(String1, Find1, Replace1)
                                            If ComboBox_color.Text = "255,0,0" Then Text1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)


                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If
                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                        String1 = Text1.TextString
                                        Dim Word1 As String = " " & Find1 & " "
                                        If InStr(String1, Word1) > 0 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = Text1.TextString
                                        Word1 = " " & Find1
                                        Dim Lungime_word1 As Integer = Len(Word1)

                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = Text1.TextString
                                        Word1 = Find1 & " "
                                        Lungime_word1 = Len(Word1)
                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        If String1 = Find1 Then
                                            Text1.UpgradeOpen()
                                            Text1.TextString = Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If
                                Next
                            Catch ex As System.Exception

                            End Try
                        End If
                        If TypeOf Ent1 Is MText Then
                            Try
                                Dim MText1 As MText = Ent1

                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                    Dim String1 As String = MText1.Contents

                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                        If InStr(String1, Find1) > 0 Then
                                            Dim String2 As String = Replace(String1, Find1, Replace1)
                                            MText1.UpgradeOpen()
                                            MText1.Contents = String2
                                            If ComboBox_color.Text = "255,0,0" Then MText1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If


                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                        String1 = MText1.Contents
                                        Dim Word1 As String = " " & Find1 & " "
                                        If InStr(String1, Word1) > 0 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = MText1.Contents
                                        Word1 = " " & Find1
                                        Dim Lungime_word1 As Integer = Len(Word1)

                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = MText1.Contents
                                        Word1 = Find1 & " "
                                        Lungime_word1 = Len(Word1)
                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        If String1 = Find1 Then
                                            MText1.UpgradeOpen()
                                            MText1.Contents = Replace1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If

                                Next
                            Catch ex As System.Exception
                            End Try

                        End If

                        If TypeOf Ent1 Is MLeader Then
                            Try
                                Dim Mleader1 As MLeader = Ent1
                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim MText1 As MText = Mleader1.MText
                                    Dim String1 As String = MText1.Contents
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)
                                    If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                        If InStr(String1, Find1) > 0 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace(String1, Find1, Replace1)
                                            If ComboBox_color.Text = "255,0,0" Then MText1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1

                                        End If
                                    End If


                                    If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                        String1 = MText1.Contents
                                        Dim Word1 As String = " " & Find1 & " "
                                        If InStr(String1, Word1) > 0 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        String1 = MText1.Contents
                                        Word1 = " " & Find1
                                        Dim Lungime_word1 As Integer = Len(Word1)

                                        If Strings.Right(String1, Lungime_word1) = Word1 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                        String1 = MText1.Contents
                                        Word1 = Find1 & " "
                                        Lungime_word1 = Len(Word1)
                                        If Strings.Left(String1, Lungime_word1) = Word1 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If

                                        If String1 = Find1 Then
                                            Mleader1.UpgradeOpen()
                                            MText1.Contents = Replace1
                                            Mleader1.MText = MText1
                                            Nr_de_obiecte = Nr_de_obiecte + 1
                                        End If
                                    End If

                                Next
                            Catch ex As System.Exception
                            End Try

                        End If

                        If TypeOf Ent1 Is BlockReference Then
                            Try
                                Dim Block1 As BlockReference = Ent1
                                For i = 0 To ListBox_FIND.Items.Count - 1
                                    Dim Find1 As String = ListBox_FIND.Items(i)
                                    Dim Replace1 As String = ListBox_REPLACE.Items(i)

                                    If Block1.AttributeCollection.Count > 0 Then
                                        For Each Atid As ObjectId In Block1.AttributeCollection
                                            Dim Atr1 As AttributeReference = Atid.GetObject(OpenMode.ForRead)
                                            If Atr1.IsMTextAttribute = True Then
                                                Dim String1 As String = Atr1.MTextAttribute.Contents
                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                    If InStr(String1, Find1) > 0 Then
                                                        Atr1.UpgradeOpen()

                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        String1 = Atr1.MTextAttribute.Contents
                                                        Mtext1.Contents = Replace(String1, Find1, Replace1)
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If

                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                    String1 = Atr1.MTextAttribute.Contents
                                                    Dim Word1 As String = " " & Find1 & " "
                                                    If InStr(String1, Word1) > 0 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Replace(String1, Word1, " " & Replace1 & " ")
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    String1 = Atr1.MTextAttribute.Contents
                                                    Word1 = " " & Find1
                                                    Dim Lungime_word1 As Integer = Len(Word1)

                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    String1 = Atr1.MTextAttribute.Contents
                                                    Word1 = Find1 & " "
                                                    Lungime_word1 = Len(Word1)
                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    If String1 = Find1 Then
                                                        Atr1.UpgradeOpen()
                                                        Dim Mtext1 As MText = Atr1.MTextAttribute
                                                        Mtext1.Contents = Replace1
                                                        Atr1.MTextAttribute = Mtext1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If
                                            Else
                                                Dim String1 As String = Atr1.TextString
                                                If CheckBox_REPLACE_ONLY_WORD.Checked = False Then
                                                    If InStr(String1, Find1) > 0 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace(String1, Find1, Replace1)
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If

                                                If CheckBox_REPLACE_ONLY_WORD.Checked = True Then
                                                    String1 = Atr1.TextString
                                                    Dim Word1 As String = " " & Find1 & " "
                                                    If InStr(String1, Word1) > 0 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace(String1, Word1, " " & Replace1 & " ")
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If

                                                    String1 = Atr1.TextString
                                                    Word1 = " " & Find1
                                                    Dim Lungime_word1 As Integer = Len(Word1)
                                                    If Strings.Right(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Strings.Left(String1, Len(String1) - Lungime_word1) & " " & Replace1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                    String1 = Atr1.TextString
                                                    Word1 = Find1 & " "
                                                    Lungime_word1 = Len(Word1)
                                                    If Strings.Left(String1, Lungime_word1) = Word1 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace1 & " " & Strings.Right(String1, Len(String1) - Lungime_word1)
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                    If String1 = Find1 Then
                                                        Atr1.UpgradeOpen()
                                                        Atr1.TextString = Replace1
                                                        Nr_de_obiecte = Nr_de_obiecte + 1
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If
                                Next
                            Catch ex As System.Exception
                            End Try
                        End If
                    Next


                    Trans1.Commit()
                End Using ' asta e de la tranzactie
            End Using
            Show_butoanele()
            MsgBox("Finished" & vbCrLf & Nr_de_obiecte & " items replaced")


            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            Show_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub ListBox_FIND_Click(sender As Object, e As EventArgs) Handles ListBox_FIND.Click
        Try
            Dim curent_index As Integer = ListBox_FIND.SelectedIndex
            If curent_index >= 0 Then
                If ListBox_FIND.Items.Count > 0 Then
                    Dim Rezultat_msg As MsgBoxResult = MsgBox("Add?", vbYesNo)
                    If Rezultat_msg = vbYes Then
                        Dim Find1 As String = InputBox("Add the text to be searched:")
                        If Not Find1 = "" Then
                            Dim Replace1 As String = InputBox("Add the replacement text:")
                            If Not Replace1 = "" Then
                                ListBox_FIND.Items.Add(Find1)
                                ListBox_REPLACE.Items.Add(Replace1)
                            End If
                        End If
                    ElseIf Rezultat_msg = vbNo Then
                        If MsgBox("Delete?", vbYesNo) = vbYes Then
                            ListBox_FIND.Items.RemoveAt(curent_index)
                            ListBox_REPLACE.Items.RemoveAt(curent_index)
                        Else
                            Dim Find1 As String = InputBox("Specify new text:")
                            If Not Find1 = "" Then
                                ListBox_FIND.Items(curent_index) = Find1
                            End If
                        End If
                    End If
                End If
            Else
                Dim Rezultat_msg As MsgBoxResult = MsgBox("Add?", vbYesNo)
                If Rezultat_msg = vbYes Then
                    Dim Find1 As String = InputBox("Add the text to be searched:")
                    If Not Find1 = "" Then
                        Dim Replace1 As String = InputBox("Add the replacement text:")
                        If Not Replace1 = "" Then
                            ListBox_FIND.Items.Add(Find1)
                            ListBox_REPLACE.Items.Add(Replace1)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ListBox_REPLACE_Click(sender As Object, e As EventArgs) Handles ListBox_REPLACE.Click
        Try
            Dim curent_index As Integer = ListBox_REPLACE.SelectedIndex
            If curent_index >= 0 Then
                If ListBox_REPLACE.Items.Count > 0 Then
                    Dim Rezultat_msg As MsgBoxResult = MsgBox("Add?", vbYesNo)
                    If Rezultat_msg = vbYes Then
                        Dim Find1 As String = InputBox("Add the text to be searched:")
                        If Not Find1 = "" Then
                            Dim Replace1 As String = InputBox("Add the replacement text:")
                            If Not Replace1 = "" Then
                                ListBox_FIND.Items.Add(Find1)
                                ListBox_REPLACE.Items.Add(Replace1)
                            End If
                        End If
                    ElseIf Rezultat_msg = vbNo Then
                        If MsgBox("Delete?", vbYesNo) = vbYes Then
                            ListBox_FIND.Items.RemoveAt(curent_index)
                            ListBox_REPLACE.Items.RemoveAt(curent_index)
                        Else
                            Dim Replace1 As String = InputBox("Specify replacement text:")
                            If Not Replace1 = "" Then
                                ListBox_REPLACE.Items(curent_index) = Replace1
                            End If
                        End If
                    End If
                End If
            Else
                Dim Rezultat_msg As MsgBoxResult = MsgBox("Add?", vbYesNo)
                If Rezultat_msg = vbYes Then
                    Dim Find1 As String = InputBox("Add the text to be searched:")
                    If Not Find1 = "" Then
                        Dim Replace1 As String = InputBox("Add the replacement text:")
                        If Not Replace1 = "" Then
                            ListBox_FIND.Items.Add(Find1)
                            ListBox_REPLACE.Items.Add(Replace1)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button_rename_files_Click(sender As Object, e As EventArgs) Handles Button_rename_files.Click
        Try
            If TextBox_destination_folder.Text = "" Then
                MsgBox("Please specify the destination folder!")
                Exit Sub
            End If

            Dim Folder_destinatie As String = TextBox_destination_folder.Text



            If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Folder_destinatie) = False Then
                MsgBox("Please specify the destination folder!")
                Exit Sub
            End If

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()

            Ascunde_butoanele()

            Dim Colectie_nume_fisiere As New Specialized.StringCollection

            For Each File1 As String In My.Computer.FileSystem.GetFiles(Folder_destinatie, FileIO.SearchOption.SearchAllSubDirectories, "*.dwg")
                Colectie_nume_fisiere.Add(File1)
            Next
            'If System.IO.File.Exists(Director_processed & "\" & Table_data.Rows.Item(i).Item("Col_New_file_name")) = False Then
            Dim row1 As Integer = 1
            For j = 0 To Colectie_nume_fisiere.Count - 1
                Dim Fisier As String = Colectie_nume_fisiere(j)
                'W1.Cells(j + 1, 1).VALUE = Fisier
                Dim nume_fisier_vechi As String = System.IO.Path.GetFileName(Fisier)
                Dim find1 As String = System.IO.Path.GetFileNameWithoutExtension(Fisier)
                Dim nume_fisier_nou As String = ""
                For i = 0 To ListBox_FIND.Items.Count - 1
                    If find1.ToUpper = ListBox_FIND.Items(i).ToString.ToUpper Then
                        nume_fisier_nou = ListBox_REPLACE.Items(i) & ".dwg"

                    End If

                Next
                If Not nume_fisier_nou = "" Then
                    Dim Path_to_fisier As String = System.IO.Path.GetDirectoryName(Fisier) & "\"
                    If System.IO.File.Exists(Path_to_fisier & nume_fisier_nou) = False Then
                        Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Fisier, nume_fisier_nou)
                    Else
                        W1.Cells(row1, 1).text = Fisier
                        row1 = row1 + 1
                    End If
                End If


            Next ' asta e de la For i = 0 To Colectie_nume_fisiere.Count - 1

            Show_butoanele()
            MsgBox("Finished")


            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            Show_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub

End Class