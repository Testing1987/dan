Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Block_layout_insert_Form
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Block_layout_insert_Form_Load(sender As Object, e As EventArgs) Handles Me.Load, Me.Click
        Incarca_existing_Blocks_to_combobox(ComboBox_existing_blocks)
        Incarca_existing_layers_to_combobox(ComboBox_existing_layers)
    End Sub
    Private Sub Button_insert_block_Click(sender As Object, e As EventArgs) Handles Button_insert_block.Click
        If Not ComboBox_existing_blocks.Text = "" And Not ComboBox_existing_layers.Text = "" Then

            Dim Start1 As Integer
            Dim End1 As Integer

            If IsNumeric(TextBox_Layout_start.Text) = True Then
                Start1 = CInt(TextBox_Layout_start.Text)
            End If

            If IsNumeric(TextBox_Layout_end.Text) = True Then
                End1 = CInt(TextBox_Layout_end.Text)
            End If

            If End1 < 1 Or Start1 < 1 Or Start1 > End1 Then
                MsgBox("Please specify the correct layouts!")
                Exit Sub
            End If

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try
                Dim Point_rezult As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify insertion point")
                PP0.AllowNone = True
                Point_rezult = Editor1.GetPoint(PP0)
                If Not Point_rezult.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If


                Using lock1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                        Dim Layoutdict As DBDictionary = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)
                        Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
                        Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

                        If Tilemode1 = 0 Then
                            If CVport1 = 2 Then
                                Editor1.SwitchToPaperSpace()
                            End If
                        Else
                            Application.SetSystemVariable("TILEMODE", 0)
                        End If

                        For Each entry As DBDictionaryEntry In Layoutdict
                            Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead)
                            If Layout1.TabOrder >= Start1 And Layout1.TabOrder <= End1 Then
                                If Layout1.TabSelected = False Then LayoutManager1.CurrentLayout = Layout1.LayoutName
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                Dim X As Double = Point_rezult.Value.X
                                Dim y As Double = Point_rezult.Value.Y
                                Dim Colectie1 As New Specialized.StringCollection
                                Dim Colectie2 As New Specialized.StringCollection
                                InsertBlock_with_multiple_atributes("", ComboBox_existing_blocks.Text, New Point3d(X, y, 0), 1, BTrecord, ComboBox_existing_layers.Text, Colectie1, Colectie2)

                            End If

                        Next






                        Trans1.Commit()

                    End Using
                End Using



                Editor1.WriteMessage(vbLf & "Command:")
            Catch ex As Exception

                Editor1.WriteMessage(vbLf & "Command:")
                MsgBox(ex.Message)
            End Try
            afiseaza_butoanele_pentru_forms(Me, Colectie1)

        End If


    End Sub


End Class