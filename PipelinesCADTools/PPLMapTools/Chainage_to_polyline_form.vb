Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Chainage_to_polyline_form
    Dim Colectie_butoane As New Specialized.StringCollection
    Private Sub Chainage_to_polyline_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Incarca_existing_layers_to_combobox(ComboBox_BLOCK_LAYER)
        If ComboBox_BLOCK_LAYER.Items.Contains("TEXT") = True Then
            ComboBox_BLOCK_LAYER.SelectedIndex = ComboBox_BLOCK_LAYER.Items.IndexOf("TEXT")
        End If
    End Sub

    Private Sub Button_DAN_Click(sender As System.Object, e As System.EventArgs) Handles Button_DAN.Click
        Try


            If IsNumeric(TextBox_scale.Text) = False Then
                With TextBox_scale
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify block scale")

                Exit Sub
            End If
            If Val(TextBox_row_start.Text) < 1 Or IsNumeric(TextBox_row_start.Text) = False Then
                With TextBox_row_start
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify start row")

                Exit Sub
            End If
            If Val(TextBox_row_end.Text) < 1 Or IsNumeric(TextBox_row_end.Text) = False Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("Please specify end row")

                Exit Sub
            End If

            If Val(TextBox_row_end.Text) < Val(TextBox_row_start.Text) Then
                With TextBox_row_end
                    .Text = ""
                    .Focus()
                End With
                MsgBox("End row smaller than start row")

                Exit Sub
            End If
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            ascunde_butoanele_pentru_forms(Me, Colectie_butoane)
            Me.Refresh()
            Dim Lock1 As DocumentLock
            Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            Using Lock1
                



                ' Dim k As Double = 1
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                    Dim start1 As Double = CDbl(TextBox_row_start.Text)
                    Dim end1 As Double = CDbl(TextBox_row_end.Text)
                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)


                    Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                    Editor1 = ThisDrawing.Editor
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select the polyline:"

                    Object_Prompt.SingleOnly = True
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)
                    If Not Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        MsgBox("Please select a polyline")
                        afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                        Exit Sub
                    Else
                        If Not TypeOf Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead) Is Curve Then
                            MsgBox("Please select a polyline")
                            afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                            Exit Sub

                        End If
                    End If


                    Dim Poly1 As Curve = Rezultat1.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)

                    For i = start1 To end1
                        Dim KP1 As Double
                        ' de aici
                        If IsNumeric(W1.Range(TextBox_kp.Text.ToUpper & i).Value) = True Then
                            KP1 = CDbl(W1.Range(TextBox_kp.Text.ToUpper & i).Value)
                        Else
                            MsgBox("See cell " & TextBox_kp.Text & i)
                            W1.Range(TextBox_kp.Text.ToUpper & i).Select()
                            afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
                            Exit Sub
                        End If



                        Dim Block_name_string As String = W1.Range(TextBox_Block_name.Text.ToUpper & i).Value

                        Dim Atr_field_name1 As String = ""
                        Dim Atr_field_value1 As String = ""

                        Dim Atr_field_name2 As String = ""
                        Dim Atr_field_value2 As String = ""

                        Dim Atr_field_name3 As String = ""
                        Dim Atr_field_value3 As String = ""

                        If Not TextBox_attribute_name1.Text = "" And Not TextBox_Value1.Text = "" Then
                            Atr_field_name1 = W1.Range(TextBox_attribute_name1.Text.ToUpper & i).Value
                            Atr_field_value1 = W1.Range(TextBox_Value1.Text.ToUpper & i).Value
                        End If

                        If Not TextBox_attribute_name2.Text = "" And Not TextBox_Value2.Text = "" Then
                            Atr_field_name2 = W1.Range(TextBox_attribute_name2.Text.ToUpper & i).Value
                            Atr_field_value2 = W1.Range(TextBox_Value2.Text.ToUpper & i).Value
                        End If
                        If Not TextBox_attribute_name3.Text = "" And Not TextBox_Value3.Text = "" Then
                            Atr_field_name3 = W1.Range(TextBox_attribute_name3.Text.ToUpper & i).Value
                            Atr_field_value3 = W1.Range(TextBox_Value3.Text.ToUpper & i).Value
                        End If

                        'de aici sunt bLOCURI
                        Dim Colectie_atr_name As New Specialized.StringCollection
                        Dim Colectie_atr_value As New Specialized.StringCollection

                        If Not Atr_field_name1 = "" And Not Atr_field_value1 = "" Then
                            Colectie_atr_name.Add(Atr_field_name1)
                            Colectie_atr_value.Add(Atr_field_value1)
                        End If
                        If Not Atr_field_name2 = "" And Not Atr_field_value2 = "" Then
                            Colectie_atr_name.Add(Atr_field_name2)
                            Colectie_atr_value.Add(Atr_field_value2)
                        End If
                        If Not Atr_field_name3 = "" And Not Atr_field_value3 = "" Then
                            Colectie_atr_name.Add(Atr_field_name3)
                            Colectie_atr_value.Add(Atr_field_value3)
                        End If

                        If KP1 <= Poly1.GetDistAtPoint(Poly1.EndPoint) Then
                            Dim x, y, Z As Double
                            If CheckBox_CSF.Checked = False Then
                                x = Poly1.GetPointAtDist(KP1).TransformBy(UCS_CURENT).X
                                y = Poly1.GetPointAtDist(KP1).TransformBy(UCS_CURENT).Y
                                Z = Poly1.GetPointAtDist(KP1).TransformBy(UCS_CURENT).Z
                            Else
                                Dim point1 As Point3d = get_csf_point_from_chainage(KP1, Poly1).TransformBy(UCS_CURENT)
                                x = point1.X
                                y = point1.Y
                                Z = point1.Z

                            End If




                            Dim Scale1 As Double = CDbl(TextBox_scale.Text)

                            InsertBlock_with_multiple_atributes(Block_name_string & ".dwg", Block_name_string, New Point3d(x, y, Z), Scale1, BTrecord, ComboBox_BLOCK_LAYER.Text, Colectie_atr_name, Colectie_atr_value)
                        Else
                            W1.Range(TextBox_kp.Text.ToUpper & i).Interior.ColorIndex = 5
                            MsgBox("Chainage bigger than length on cell " & TextBox_kp.Text.ToUpper & i)
                        End If
                        'asta e de la INSERT BLOCKS
                    Next





                    Trans1.Commit()
                    ' asta e de la tranzactie
                End Using


                afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)


                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                ' asta e de la lock
            End Using



        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie_butoane)
            MsgBox(ex.Message)
        End Try
    End Sub
   

    Private Sub Panel_COLUMNS_Click(sender As Object, e As EventArgs) Handles Panel_COLUMNS.Click
        Incarca_existing_layers_to_combobox(ComboBox_BLOCK_LAYER)
        If ComboBox_BLOCK_LAYER.Items.Contains("TEXT") = True Then
            ComboBox_BLOCK_LAYER.SelectedIndex = ComboBox_BLOCK_LAYER.Items.IndexOf("TEXT")
        End If
    End Sub


End Class