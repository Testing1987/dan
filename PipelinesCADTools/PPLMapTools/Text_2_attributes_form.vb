Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Text_2_attributes_form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Data_table_block_atr As System.Data.DataTable
    Dim Data_table_text As System.Data.DataTable
    Dim DWG_Block As String = ""
    Dim Index_radio_checked As Integer = -1

    Private Sub Text_2_attributes_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Button_load_text.Visible = False
        Button_insert_block.Visible = False
        Incarca_existing_layers_to_combobox(ComboBox_layers)
        If ComboBox_layers.Items.Contains("TEXT") = True Then
            ComboBox_layers.SelectedIndex = ComboBox_layers.Items.IndexOf("TEXT")
        End If
    End Sub
    Private Sub Panel_layers_Click(sender As Object, e As EventArgs) Handles Panel_layers.Click
        Incarca_existing_layers_to_combobox(ComboBox_layers)
        If ComboBox_layers.Items.Contains("TEXT") = True Then
            ComboBox_layers.SelectedIndex = ComboBox_layers.Items.IndexOf("TEXT")
        End If
    End Sub
    Private Sub Panel_blocks_Click(sender As Object, e As EventArgs) Handles Panel_BLOCK.Click
        For Each Control1 In Panel_BLOCK.Controls
            If TypeOf Control1 Is Windows.Forms.RadioButton Then
                Dim Radio1 As Windows.Forms.RadioButton = Control1
                Radio1.Checked = False
            End If
        Next
    End Sub

    Private Sub Button_LOAD_BLOCK_Click(sender As Object, e As EventArgs) Handles Button_LOAD_BLOCK.Click

        Try

            Dim filedialog1 As New System.Windows.Forms.OpenFileDialog
            filedialog1.InitialDirectory = "G:\_HMM\Design\Programming\Autocad Blocks"
            filedialog1.Filter = "dwg files (*.dwg)|*.dwg|All files (*.*)|*.*"
            If filedialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

                Panel_BLOCK.Controls.Clear()

                DWG_Block = filedialog1.FileName
                Label_DWG_BLOCK.Text = DWG_Block

                Dim Database1 As New Database(False, True)
                Database1.ReadDwgFile(DWG_Block, IO.FileShare.Read, False, Nothing)
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                    Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(Autodesk.AutoCAD.DatabaseServices.BlockTableRecord.ModelSpace), OpenMode.ForRead)

                    Data_table_block_atr = New System.Data.DataTable
                    Data_table_block_atr.Columns.Add("TAG", GetType(String))
                    Data_table_block_atr.Columns.Add("X", GetType(Double))
                    Data_table_block_atr.Columns.Add("Y", GetType(Double))
                    Dim Index_data_table As Integer = 0


                    Button_load_text.Visible = True

                    For Each ID1 As ObjectId In BTrecord
                        Dim Ent1 As Entity = Trans1.GetObject(ID1, OpenMode.ForRead)
                        If TypeOf Ent1 Is AttributeDefinition Then
                            Dim Atr1 As AttributeDefinition = Ent1
                            Data_table_block_atr.Rows.Add()
                            Data_table_block_atr.Rows(Index_data_table).Item("X") = Atr1.Position.X
                            Data_table_block_atr.Rows(Index_data_table).Item("Y") = Atr1.Position.Y
                            Data_table_block_atr.Rows(Index_data_table).Item("TAG") = Atr1.Tag
                            Index_data_table = Index_data_table + 1
                        End If


                    Next
                    '" DESC,"  " ASC"
                    Data_table_block_atr = Sort_data_table_2_columns(Data_table_block_atr, "Y", " DESC,", "X", " ASC")

                    For i = 0 To Data_table_block_atr.Rows.Count - 1
                        Dim Combo1 As New Windows.Forms.ComboBox
                        Combo1.Text = Data_table_block_atr.Rows(i).Item("TAG")
                        Dim myfont As New System.Drawing.Font("Arial", 9, Drawing.FontStyle.Bold)
                        Combo1.Font = myfont
                        Combo1.Size = New System.Drawing.Size(130, 23)
                        Combo1.Location = New System.Drawing.Point(5, 5 + (Combo1.Height + 5) * i)
                        Combo1.Name = "Combobox1_" & i
                        Combo1.DropDownStyle = Windows.Forms.ComboBoxStyle.DropDown

                        Panel_BLOCK.Controls.Add(Combo1)


                        Dim Radio1 As New Windows.Forms.RadioButton
                        Radio1.Text = ""
                        Radio1.Font = myfont
                        Radio1.Size = New System.Drawing.Size(14, 13)
                        Radio1.Location = New System.Drawing.Point(142, Combo1.Location.Y + Combo1.Height / 4)
                        Radio1.Name = "Radio1_" & i
                        Radio1.Checked = False
                        Panel_BLOCK.Controls.Add(Radio1)

                        If i < 21 Then
                            Panel_BLOCK.Height = Combo1.Location.Y + Combo1.Height + 10
                        End If

                        If Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5 < Panel_text.Location.Y + Panel_text.Height + 5 Then
                            Button_LOAD_BLOCK.Location = New System.Drawing.Point(5, Panel_text.Location.Y + Panel_text.Height + 5)
                            Button_load_text.Location = New System.Drawing.Point(162, Panel_text.Location.Y + Panel_text.Height + 5)
                            Button_insert_block.Location = New System.Drawing.Point(305, Panel_text.Location.Y + Panel_text.Height + 5)
                            Button_reddfine_block.Location = New System.Drawing.Point(450, Panel_text.Location.Y + Panel_text.Height + 5)
                        Else
                            Button_LOAD_BLOCK.Location = New System.Drawing.Point(5, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                            Button_load_text.Location = New System.Drawing.Point(162, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                            Button_insert_block.Location = New System.Drawing.Point(305, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                            Button_reddfine_block.Location = New System.Drawing.Point(450, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                        End If
                        Me.Height = Button_LOAD_BLOCK.Location.Y + Button_LOAD_BLOCK.Height + 40

                    Next
                End Using

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Private Sub Button_load_text_Click(sender As Object, e As EventArgs) Handles Button_load_text.Click
        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        For Each Control1 As Windows.Forms.Control In Panel_BLOCK.Controls
            If Control1.Location.X = 160 Then
                Panel_BLOCK.Controls.Remove(Control1)
            End If
        Next

        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                    Panel_text.Controls.Clear()
                    Data_table_text = New System.Data.DataTable
                    Data_table_text.Columns.Add("TEXT", GetType(String))
                    Data_table_text.Columns.Add("X", GetType(Double))
                    Data_table_text.Columns.Add("Y", GetType(Double))
                    Dim Index_data_table As Integer = 0

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select the text"

                    Object_Prompt.SingleOnly = False
                    Rezultat1 = ThisDrawing.Editor.GetSelection(Object_Prompt)


                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If


                    For i = 0 To Rezultat1.Value.Count - 1
                        Dim ent1 As Entity = Rezultat1.Value.Item(i).ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf ent1 Is DBText Then
                            Dim Text1 As DBText = ent1
                            Data_table_text.Rows.Add()
                            Data_table_text.Rows(Index_data_table).Item("X") = Text1.Position.X
                            Data_table_text.Rows(Index_data_table).Item("Y") = Text1.Position.Y
                            Data_table_text.Rows(Index_data_table).Item("TEXT") = Text1.TextString


                            Index_data_table = Index_data_table + 1
                        End If
                        If TypeOf ent1 Is MText Then
                            Dim Text1 As MText = ent1
                            Data_table_text.Rows.Add()
                            Data_table_text.Rows(Index_data_table).Item("X") = Text1.Location.X
                            Data_table_text.Rows(Index_data_table).Item("Y") = Text1.Location.Y
                            Data_table_text.Rows(Index_data_table).Item("TEXT") = Text1.Text

                            Index_data_table = Index_data_table + 1
                        End If

                    Next



                    Data_table_text = Sort_data_table_2_columns(Data_table_text, "Y", " DESC,", "X", " ASC")

                    If Data_table_text.Rows.Count < Data_table_block_atr.Rows.Count Then
                        For i = 1 To Data_table_block_atr.Rows.Count - Data_table_text.Rows.Count
                            Data_table_text.Rows.Add()
                        Next
                    End If

                    For i = 0 To Data_table_text.Rows.Count - 1

                        Dim Combo2 As New Windows.Forms.ComboBox
                        If IsDBNull(Data_table_text.Rows(i).Item("TEXT")) = False Then Combo2.Text = Data_table_text.Rows(i).Item("TEXT")
                        Dim myfont As New System.Drawing.Font("Arial", 9, Drawing.FontStyle.Bold)
                        Combo2.Font = myfont
                        Combo2.Size = New System.Drawing.Size(285, 23)
                        Combo2.Location = New System.Drawing.Point(5, 5 + (Combo2.Height + 5) * i)
                        Combo2.Name = "Combobox2_" & i
                        Combo2.DropDownStyle = Windows.Forms.ComboBoxStyle.DropDown
                        Panel_text.Controls.Add(Combo2)

                        Dim Radio2 As New Windows.Forms.RadioButton
                        Radio2.Text = ""
                        Radio2.Font = myfont
                        Radio2.Size = New System.Drawing.Size(14, 13)
                        Radio2.Location = New System.Drawing.Point(296, Combo2.Location.Y + Combo2.Height / 4)
                        Radio2.Name = "Radio2_" & i
                        Radio2.Checked = False
                        Panel_text.Controls.Add(Radio2)
                        AddHandler Radio2.CheckedChanged, AddressOf Radio2_CheckedChanged

                        If i < 21 Then
                            Panel_text.Height = Combo2.Location.Y + Combo2.Height + 10
                        End If

                        If Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5 < Panel_text.Location.Y + Panel_text.Height + 5 Then
                            Button_LOAD_BLOCK.Location = New System.Drawing.Point(5, Panel_text.Location.Y + Panel_text.Height + 5)
                            Button_load_text.Location = New System.Drawing.Point(162, Panel_text.Location.Y + Panel_text.Height + 5)
                            Button_insert_block.Location = New System.Drawing.Point(305, Panel_text.Location.Y + Panel_text.Height + 5)
                            Button_reddfine_block.Location = New System.Drawing.Point(450, Panel_text.Location.Y + Panel_text.Height + 5)
                        Else
                            Button_LOAD_BLOCK.Location = New System.Drawing.Point(5, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                            Button_load_text.Location = New System.Drawing.Point(162, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                            Button_insert_block.Location = New System.Drawing.Point(305, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                            Button_reddfine_block.Location = New System.Drawing.Point(450, Panel_BLOCK.Location.Y + Panel_BLOCK.Height + 5)
                        End If

                        Me.Height = Button_LOAD_BLOCK.Location.Y + Button_LOAD_BLOCK.Height + 40

                    Next
                    Button_insert_block.Visible = True
                    Trans1.Commit()


                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Radio2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.Checked = True Then
            Dim Nume1 As String = sender.Name
            Dim _poz1 As Integer = InStr(Nume1, "_")
            Nume1 = Strings.Right(Nume1, Len(Nume1) - _poz1)
            Dim Nr1 As Integer = -1
            Dim Nume2 As String
            Dim Nr2 As Integer = -1
            Dim Radiob1 As Windows.Forms.RadioButton

            If IsNumeric(Nume1) = True Then
                Nr1 = CInt(Nume1)
                For Each Control1 In Panel_BLOCK.Controls
                    If TypeOf Control1 Is Windows.Forms.RadioButton Then
                        Dim Radio1 As Windows.Forms.RadioButton = Control1
                        If Radio1.Checked = True Then
                            Nume2 = Radio1.Name
                            Dim _poz2 As Integer = InStr(Nume2, "_")
                            Nume2 = Strings.Right(Nume2, Len(Nume2) - _poz2)
                            If IsNumeric(Nume2) = True Then
                                Nr2 = CInt(Nume2)
                                Radiob1 = Radio1
                            End If
                        End If


                    End If

                Next



            End If

            If Not Nr1 = -1 And Not Nr2 = -1 Then
                If IsDBNull(Data_table_text.Rows(Nr1).Item("TEXT")) = True Then
                    Data_table_text.Rows(Nr1).Item("TEXT") = " "
                End If
                If IsDBNull(Data_table_text.Rows(Nr2).Item("TEXT")) = True Then
                    Data_table_text.Rows(Nr2).Item("TEXT") = " "
                End If
                Dim Temp1 As String
                Temp1 = Data_table_text.Rows(Nr1).Item("TEXT")
                Data_table_text.Rows(Nr1).Item("TEXT") = Data_table_text.Rows(Nr2).Item("TEXT")
                Data_table_text.Rows(Nr2).Item("TEXT") = Temp1
                If IsDBNull(Data_table_text.Rows(Nr1).Item("X")) = True Then
                    Data_table_text.Rows(Nr1).Item("X") = -1
                End If
                If IsDBNull(Data_table_text.Rows(Nr2).Item("X")) = True Then
                    Data_table_text.Rows(Nr2).Item("X") = -1
                End If
                If IsDBNull(Data_table_text.Rows(Nr1).Item("Y")) = True Then
                    Data_table_text.Rows(Nr1).Item("Y") = -1
                End If
                If IsDBNull(Data_table_text.Rows(Nr2).Item("Y")) = True Then
                    Data_table_text.Rows(Nr2).Item("Y") = -1
                End If
                Dim TempX As Double
                TempX = Data_table_text.Rows(Nr1).Item("X")
                Data_table_text.Rows(Nr1).Item("X") = Data_table_text.Rows(Nr2).Item("X")
                Data_table_text.Rows(Nr2).Item("X") = TempX
                Dim TempY As Double
                TempY = Data_table_text.Rows(Nr1).Item("Y")
                Data_table_text.Rows(Nr1).Item("Y") = Data_table_text.Rows(Nr2).Item("Y")
                Data_table_text.Rows(Nr2).Item("Y") = TempY
              


                Dim Combo_box1 As Windows.Forms.ComboBox
                Dim Combo_box2 As Windows.Forms.ComboBox
                Dim Radio_1 As Windows.Forms.RadioButton
                Dim Radio_2 As Windows.Forms.RadioButton
                For Each Control1 In Panel_text.Controls
                    If TypeOf Control1 Is Windows.Forms.ComboBox Then
                        Dim Combo1 As Windows.Forms.ComboBox = Control1
                        If Combo1.Name = "Combobox2_" & Nr1 Then
                            Combo_box1 = Combo1
                        End If
                        If Combo1.Name = "Combobox2_" & Nr2 Then
                            Combo_box2 = Combo1
                        End If

                    End If
                    If TypeOf Control1 Is Windows.Forms.RadioButton Then
                        Dim Radio1 As Windows.Forms.RadioButton = Control1
                        If Radio1.Name = "Radio2_" & Nr1 Then
                            Radio_1 = Radio1
                        End If
                        If Radio1.Name = "Radio2_" & Nr2 Then
                            Radio_2 = Radio1
                        End If


                    End If
                Next
                If IsNothing(Combo_box1) = False And IsNothing(Combo_box2) = False Then
                    Dim Tempor1 As String = Combo_box1.Text
                    Combo_box1.Text = Combo_box2.Text
                    Combo_box2.Text = Tempor1
                    Radio_1.Checked = False
                    Radiob1.Checked = False
                End If

            Else
                For Each Control1 In Panel_text.Controls
                    If TypeOf Control1 Is Windows.Forms.RadioButton Then
                        Dim Radio1 As Windows.Forms.RadioButton = Control1

                    End If
                Next
                If Not Index_radio_checked = -1 Then
                    For Each Control1 In Panel_text.Controls
                        If TypeOf Control1 Is Windows.Forms.ComboBox Then
                            Dim Combo1 As Windows.Forms.ComboBox = Control1


                        End If
                    Next
                End If

            End If



        End If
    End Sub


    Private Sub Button_insert_block_Click(sender As Object, e As EventArgs) Handles Button_insert_block.Click
        Try
            If Data_table_block_atr.Rows.Count > 0 Then
                ascunde_butoanele_pentru_forms(Me, Colectie1)
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim Colectie_atr_name As New Specialized.StringCollection
                Dim Colectie_atr_value As New Specialized.StringCollection
                Dim x As Double = 0
                Dim Y As Double = 0
                Dim Scale1 As Double = 1

                If IsNumeric(TextBox_scale.Text) = True Then
                    Scale1 = CDbl(TextBox_scale.Text)
                End If
                If IsNumeric(TextBox_X.Text) = True Then
                    x = CDbl(TextBox_X.Text)
                End If
                If IsNumeric(TextBox_Y.Text) = True Then
                    Y = CDbl(TextBox_Y.Text)
                End If

                If Not DWG_Block = "" Then
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                            For Each control1 As Windows.Forms.Control In Panel_text.Controls
                                If TypeOf control1 Is Windows.Forms.ComboBox Then
                                    Dim CCC1 As Windows.Forms.ComboBox = control1

                                    Dim Index1 As String = CCC1.Name
                                    Dim _poz1 As Integer = InStr(Index1, "_")
                                    Index1 = Strings.Right(Index1, Len(Index1) - _poz1)
                                    If IsNumeric(Index1) = True Then
                                        Data_table_text.Rows(CInt(Index1)).Item("TEXT") = CCC1.Text
                                    End If


                                End If
                            Next

                            For i = 0 To Data_table_block_atr.Rows.Count - 1
                                If i <= Data_table_text.Rows.Count - 1 Then
                                    If IsDBNull(Data_table_block_atr.Rows(i).Item("TAG")) = False And IsDBNull(Data_table_text.Rows(i).Item("TEXT")) = False Then
                                        Colectie_atr_name.Add(Data_table_block_atr.Rows(i).Item("TAG"))
                                        Colectie_atr_value.Add(Data_table_text.Rows(i).Item("TEXT"))

                                    End If
                                End If
                            Next

                            Dim NUME_BLOCK As String
                            NUME_BLOCK = IO.Path.GetFileNameWithoutExtension(DWG_Block)

                            'InsertBlock_with_multiple_atributes_from_browser_location(DWG_Block, NUME_BLOCK, New Point3d(x, Y, 0), Scale1, BTrecord, ComboBox_layers.Text, Colectie_atr_name, Colectie_atr_value)
                            InsertBlock_with_multiple_atributes(DWG_Block, NUME_BLOCK, New Point3d(x, Y, 0), Scale1, BTrecord, ComboBox_layers.Text, Colectie_atr_name, Colectie_atr_value)

                            Trans1.Commit()
                        End Using
                    End Using
                End If
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

        End Try
    End Sub

    Private Sub Button_reddfine_block_Click(sender As Object, e As EventArgs) Handles Button_reddfine_block.Click
        Try
            If Data_table_block_atr.Rows.Count > 0 Then
                ascunde_butoanele_pentru_forms(Me, Colectie1)
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                If Not DWG_Block = "" Then
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                            Dim NUME_BLOCK As String
                            NUME_BLOCK = IO.Path.GetFileNameWithoutExtension(DWG_Block)
                            RedefineBlock_from_browser_location(DWG_Block, NUME_BLOCK)
                            Trans1.Commit()
                        End Using
                    End Using
                End If
            End If
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Catch ex As Exception
            MsgBox(ex.Message)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

        End Try
    End Sub
End Class

