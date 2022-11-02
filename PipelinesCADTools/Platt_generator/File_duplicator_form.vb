Imports ACSMCOMPONENTS20Lib
Public Class File_duplicator_form
    Dim Freeze_operations As Boolean = False

    Private Sub Button_sheet_set_template_Click(sender As Object, e As EventArgs) Handles Button_sheet_set_template.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim FileBrowserDialog1 As New Windows.Forms.OpenFileDialog
                FileBrowserDialog1.Filter = "Sheet Set Files (*.dst)|*.dst|All Files (*.*)|*.*"
                If FileBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim Path1 As String = FileBrowserDialog1.FileName
                    TextBox_sheet_set_template.Text = Path1
                    TextBox_NAME.Text = System.IO.Path.GetFileNameWithoutExtension(Path1)
                    Dim SheetSet_manager As New AcSmSheetSetMgr

                    Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase(Path1, False)
                    Dim sheetSet As AcSmSheetSet = SheetSet_database.GetSheetSet()
                    If LockDatabase(SheetSet_database, True) = True Then
                        Dim Location1 As IAcSmFileReference = sheetSet.GetNewSheetLocation
                        TextBox_Storage_folder.Text = Location1.GetFileName
                        Dim Publish1 As IAcSmPublishOptions = sheetSet.GetPublishOptions
                        Dim Location2 As IAcSmFileReference = Publish1.GetDefaultOutputdir
                        TextBox_publish_folder.Text = Location2.GetFileName

                        Dim fileReference As IAcSmFileReference
                        fileReference = sheetSet.GetAltPageSetups
                        If IsNothing(fileReference) = False Then
                            TextBox_page_setup_overrides.Text = fileReference.GetFileName
                        End If





                        LockDatabase(SheetSet_database, False)
                    End If
                    SheetSet_manager.Close(SheetSet_database)


                End If
            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub

    Private Sub Button_duplicate__Click(sender As Object, e As EventArgs) Handles Button_duplicate_.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Dim Start1 As Integer = -1
                If IsNumeric(TextBox_INCREMENT_START.Text) = True Then Start1 = CInt(TextBox_INCREMENT_START.Text)
                Dim End1 As Integer = -1
                If IsNumeric(TextBox_INCREMENT_END.Text) = True Then End1 = CInt(TextBox_INCREMENT_END.Text)
                If Start1 = -1 Or End1 = -1 Then
                    Freeze_operations = False
                    Exit Sub
                End If
                Dim Extra_zero As String = ""
                If Strings.Left(TextBox_INCREMENT_START.Text, 1) = "0" Then
                    Extra_zero = "0"
                End If
                If Strings.Left(TextBox_INCREMENT_START.Text, 2) = "00" Then
                    Extra_zero = "00"
                End If
                If Strings.Left(TextBox_INCREMENT_START.Text, 3) = "000" Then
                    Extra_zero = "000"
                End If

                If System.IO.File.Exists(TextBox_sheet_set_template.Text) = True Then
                    Dim File1 As String = TextBox_sheet_set_template.Text

                    Dim SheetSet_manager As New AcSmSheetSetMgr




                    Dim Path1 As String = System.IO.Path.GetDirectoryName(File1)
                    Dim Ext1 As String = System.IO.Path.GetExtension(File1)
                    Dim Name1 As String = TextBox_NAME.Text
                    For i = Start1 To End1
                        Dim suffix1 As String = ""

                        If Extra_zero = "0" Then
                            If i < 10 Then
                                suffix1 = Extra_zero
                            End If
                        End If

                        If Extra_zero = "00" Then
                            If i < 10 Then
                                suffix1 = Extra_zero
                            ElseIf i < 100 And i > 9 Then
                                suffix1 = "0"
                            End If
                        End If


                        If Extra_zero = "000" Then
                            If i < 10 Then
                                suffix1 = Extra_zero
                            ElseIf i < 100 And i > 9 Then
                                suffix1 = "00"
                            ElseIf i < 1000 And i > 99 Then
                                suffix1 = "0"
                            End If
                        End If

                        Dim File2 As String = Path1 & "\" & Name1 & "_" & suffix1 & i & Ext1
                        Dim Illegal() As Char = IO.Path.GetInvalidPathChars


                        For j = 0 To Len(Illegal) - 1
                            If File2.Contains(Illegal(j)) = True Then
                                MsgBox("Illegal characters in file name")
                                Freeze_operations = False
                                Exit Sub
                            End If
                        Next


                        If System.IO.File.Exists(File2) = False And System.IO.Directory.Exists(TextBox_Storage_folder.Text) = True And _
                                                                    System.IO.Directory.Exists(TextBox_publish_folder.Text) = True And _
                                                                    System.IO.File.Exists(TextBox_page_setup_overrides.Text) = True Then

                            System.IO.File.Copy(File1, File2)
                            Dim Sheetset_name As String = System.IO.Path.GetFileNameWithoutExtension(File2)
                            Dim Sheetset_description As String = System.IO.Path.GetFileNameWithoutExtension(File2)




                            Dim SheetSet_database2 As AcSmDatabase = SheetSet_manager.OpenDatabase(File2, False)
                            Dim sheetSet2 As AcSmSheetSet = SheetSet_database2.GetSheetSet()


                            If LockDatabase(SheetSet_database2, True) = True Then

                                sheetSet2.SetDesc(Sheetset_description)
                                sheetSet2.SetName(Sheetset_name)

                                Dim Location1 As IAcSmFileReference = sheetSet2.GetNewSheetLocation
                                Location1.SetFileName(TextBox_Storage_folder.Text)


                                Dim Publish2 As IAcSmPublishOptions = sheetSet2.GetPublishOptions
                                Dim Location2 As IAcSmFileReference = Publish2.GetDefaultOutputdir
                                Location2.SetFileName(TextBox_publish_folder.Text)

                                Dim fileReference As New AcSmFileReference



                                ' Create a new reference to the page setup overrides to use

                                fileReference.InitNew(SheetSet_database2)

                                fileReference.SetFileName(TextBox_page_setup_overrides.Text)



                                ' Set the drawing template that contains the page setup overrides

                                sheetSet2.SetAltPageSetups(fileReference)



                                Dim EnumSheets2 As IAcSmEnumComponent = sheetSet2.GetSheetEnumerator()
                                Dim smComponent2 As IAcSmComponent
                                Dim sheet2 As IAcSmSheet
                                smComponent2 = EnumSheets2.Next()
                                While True
                                    If smComponent2 Is Nothing Then
                                        Exit While
                                    End If
                                    sheet2 = TryCast(smComponent2, IAcSmSheet)
                                    If IsNothing(sheet2) = False Then
                                        sheetSet2.RemoveSheet(sheet2)
                                    End If
                                    smComponent2 = EnumSheets2.Next()
                                End While


                                LockDatabase(SheetSet_database2, False)
                            End If
                            SheetSet_manager.Close(SheetSet_database2)
                        Else
                            MsgBox("File " & vbCrLf & File2 & vbCrLf & " exists!" & vbCrLf & _
                                   "OR" & vbCrLf & _
                                   TextBox_Storage_folder.Text & " does not exist!" & vbCrLf _
                                   & "OR" & vbCrLf & _
                                   TextBox_publish_folder.Text & " does not exist!" & vbCrLf _
                                   & "OR" & vbCrLf & _
                                   TextBox_page_setup_overrides.Text & " does not exist!")

                            Freeze_operations = False
                            Exit Sub
                        End If
                    Next
                    MsgBox("File duplicated")

                Else
                    MsgBox("File does not exist!")
                End If

            Catch ex As System.Exception
                Freeze_operations = False
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            Freeze_operations = False
        End If
    End Sub
End Class