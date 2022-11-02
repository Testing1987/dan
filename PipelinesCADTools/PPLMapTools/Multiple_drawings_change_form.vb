Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Multiple_drawings_change_form

    Dim Freeze_operations As Boolean = False
    Private Sub Multiple_drawings_change_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim Descr_blocK_Modify As String
        Descr_blocK_Modify = "It looks for the layer[Clouds]. If found Creates a copy of the layer and names it [CLOUDS (OLD)]" & _
            vbCrLf & "It looks for the block named [REV_TRI] and [TITLEBLOCK NOTE-XREF]" & vbCrLf & _
            "If found in the block definition is changing objects layer from [Clouds] to [CLOUDS (OLD)]"


        ToolTip1.SetToolTip(Button_block_modify, Descr_blocK_Modify)
        ToolTip1.SetToolTip(Button_load_DWG, "select the drawings you wanted edited")

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

    Private Sub Button_block_modify_Click(sender As Object, e As EventArgs) Handles Button_block_modify.Click

        If Freeze_operations = False Then
            Freeze_operations = True

            Try
                Try


                    If ListBox_DWG.Items.Count > 0 Then
                        Using lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                            Using Trans11 As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                                Dim Colectie_facute As New Specialized.StringCollection


                                For i = 0 To ListBox_DWG.Items.Count - 1
                                    Dim Drawing1 As String = ListBox_DWG.Items(i)
                                    Dim Changed As Boolean = False

                                    If IO.File.Exists(Drawing1) = True Then
                                        Dim Database1 As New Database(False, True)
                                        Database1.ReadDwgFile(Drawing1, IO.FileShare.ReadWrite, True, Nothing)
                                        HostApplicationServices.WorkingDatabase = Database1


                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                                            Dim LayerTable1 As LayerTable = Trans1.GetObject(Database1.LayerTableId, OpenMode.ForWrite)
                                            If LayerTable1.Has("Clouds") = True Then
                                                Dim layer1 As LayerTableRecord = Trans1.GetObject(LayerTable1("Clouds"), OpenMode.ForRead)
                                                Dim Layer_clone As New LayerTableRecord
                                                Layer_clone = layer1.Clone
                                                Layer_clone.Name = "CLOUDS (OLD)"
                                                If LayerTable1.Has("CLOUDS (OLD)") = False Then
                                                    LayerTable1.Add(Layer_clone)
                                                    Trans1.AddNewlyCreatedDBObject(Layer_clone, True)
                                                End If
                                            End If
                                            Trans1.Commit()
                                        End Using


                                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                                            Dim BlockTable1 As BlockTable = Trans1.GetObject(Database1.BlockTableId, OpenMode.ForRead)
                                            Dim ID_block As ObjectId = ObjectId.Null

                                            Dim LayerTable1 As LayerTable = Trans1.GetObject(Database1.LayerTableId, OpenMode.ForWrite)
                                            If LayerTable1.Has("Clouds") = True Then

                                                If BlockTable1.Has("TITLEBLOCK NOTE-XREF") = True Then
                                                    ID_block = BlockTable1("TITLEBLOCK NOTE-XREF")

                                                    Dim btREC As BlockTableRecord = Trans1.GetObject(ID_block, OpenMode.ForWrite)
                                                    For Each ID1 In btREC
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(ID1, OpenMode.ForWrite), Entity)
                                                        If Not Ent1 = Nothing Then
                                                            If Ent1.Layer = "Clouds" Then
                                                                Ent1.Layer = "CLOUDS (OLD)"
                                                                Changed = True
                                                            End If



                                                        End If



                                                    Next



                                                End If



                                                If BlockTable1.Has("REV_TRI") = True Then
                                                    ID_block = BlockTable1("REV_TRI")

                                                    Dim btREC As BlockTableRecord = Trans1.GetObject(ID_block, OpenMode.ForWrite)
                                                    For Each ID1 In btREC
                                                        Dim Ent1 As Entity = TryCast(Trans1.GetObject(ID1, OpenMode.ForWrite), Entity)
                                                        If Not Ent1 = Nothing Then
                                                            If Ent1.Layer = "Clouds" Then
                                                                Ent1.Layer = "CLOUDS (OLD)"
                                                                Changed = True
                                                            End If



                                                        End If



                                                    Next



                                                End If

                                            End If






                                            Trans1.Commit()


                                        End Using


                                        Database1.SaveAs(Drawing1, True, DwgVersion.Current, Database1.SecurityParameters)
                                        Database1.Dispose()
                                        HostApplicationServices.WorkingDatabase = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
                                        If Changed = True Then
                                            Colectie_facute.Add(Drawing1)
                                        End If

                                    End If


                                Next

                                Trans11.Commit()

                                If Colectie_facute.Count > 0 Then
                                    For Each Drawing1 As String In Colectie_facute
                                        ListBox_DWG.Items.Remove(Drawing1)
                                    Next
                                End If

                            End Using
                        End Using
                    End If





                Catch ex As System.SystemException
                    MsgBox(ex.Message)

                End Try

            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
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


    Private Sub Button_read_Xref_Click(sender As Object, e As EventArgs) Handles Button_read_Xref.Click
        If Freeze_operations = False Then
            Freeze_operations = True
            Try
                Using lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                    Using Trans1 As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                        Dim XRgraph As XrefGraph = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.GetHostDwgXrefGraph(False)


                        Panel_xref_name.Controls.Clear()
                        Panel_xref_path.Controls.Clear()


                        Dim Y_index_panel As Integer = 0
                        For i = 1 To XRgraph.NumNodes - 1 ' la zero ai thisdrawing
                            Dim XrNode As XrefGraphNode = XRgraph.GetXrefNode(i)
                            If XrNode.IsNested = False Then
                                Dim BlockTableRec1 As BlockTableRecord = Trans1.GetObject(XrNode.BlockTableRecordId, OpenMode.ForRead)
                                Dim Fisier_with_path As String = BlockTableRec1.PathName
                                Dim Fisier As String = ""
                                Dim Path As String = ""

                                If System.IO.File.Exists(Fisier_with_path) = True Then
                                    Fisier = System.IO.Path.GetFileName(Fisier_with_path)
                                    Path = System.IO.Path.GetDirectoryName(Fisier_with_path)
                                    If Strings.Right(Path, 1) <> "\" Then Path = Path & "\"
                                Else
                                    Dim poz As Integer = Strings.InStrRev(Fisier_with_path, "\")
                                    Fisier = Strings.Mid(Fisier_with_path, poz + 1)
                                End If

                                Dim TextBox11 As New Windows.Forms.TextBox
                                TextBox11.Font = New System.Drawing.Font("Arial", 9.75, System.Drawing.FontStyle.Bold)
                                TextBox11.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                TextBox11.Size = New System.Drawing.Size(188, 22)
                                TextBox11.Text = Fisier
                                TextBox11.BackColor = Drawing.Color.White
                                TextBox11.ForeColor = Drawing.Color.Black
                                Panel_xref_name.Controls.Add(TextBox11)

                                Dim TextBox22 As New Windows.Forms.TextBox
                                TextBox22.Font = New System.Drawing.Font("Arial", 9.75, System.Drawing.FontStyle.Bold)
                                TextBox22.Location = New System.Drawing.Point(3, 3 + 27 * Y_index_panel)
                                TextBox22.Size = New System.Drawing.Size(540, 22)
                                TextBox22.Text = Path
                                TextBox22.BackColor = Drawing.Color.White
                                TextBox22.ForeColor = Drawing.Color.Black
                                Panel_xref_path.Controls.Add(TextBox22)

                                Dim CheckBox22 As New Windows.Forms.CheckBox
                                CheckBox22.Font = New System.Drawing.Font("Arial", 9.75, System.Drawing.FontStyle.Bold)
                                CheckBox22.Location = New System.Drawing.Point(553, 3 + 27 * Y_index_panel)
                                Panel_xref_path.Controls.Add(CheckBox22)


                                Y_index_panel = Y_index_panel + 1
                            End If
                        Next

                        Trans1.Commit()
                    End Using
                End Using

            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command: DONE!")
            Freeze_operations = False
        End If
    End Sub
End Class