Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Dimension_change_form
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Dimension_change_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        ComboBox_decimals.SelectedIndex = 0
    End Sub
    Private Sub Button_change_Click(sender As System.Object, e As System.EventArgs) Handles Button_change.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            If IsNumeric(TextBox_viewport_Scale.Text) = False Then
                MsgBox("NOT NUMERIC viewport scale")
                Exit Sub
            End If

            If IsNumeric(TextBox_SCALE_PS.Text) = False Then
                MsgBox("NOT NUMERIC Scale")
                Exit Sub
            End If
            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Rezultat1 = Editor1.SelectImplied

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                Else
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = "Select objects:"

                    Object_Prompt.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)
                End If





                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        ascunde_butoanele_pentru_forms(Me, Colectie1)

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                            For i = 1 To Rezultat1.Value.Count

                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i - 1)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Dimension Then
                                    Ent1.UpgradeOpen()
                                    Dim Dim1 As Dimension = Ent1
                                    Dim1.Dimdec = CInt(ComboBox_decimals.Text)


                                    If CheckBox_Scale_linear.Checked = True Then
                                        If RadioButton_mm.Checked = True Then
                                            Dim1.Dimlfac = -1 * CDbl(TextBox_SCALE_PS.Text)
                                        Else
                                            Dim1.Dimlfac = -1 * CDbl(TextBox_SCALE_PS.Text) / 1000
                                        End If
                                    End If



                                End If






                            Next

                            Trans1.Commit()

                        End Using

                    Else

                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox_viewport_Scale_Click(sender As Object, e As System.EventArgs) Handles TextBox_viewport_Scale.Click
        Try
            If IsNumeric(TextBox_viewport_Scale.Text) = True And IsNumeric(TextBox_printScale.Text) = True Then
                Dim viewport_scale As Double = CDbl(TextBox_viewport_Scale.Text)
                If RadioButton_mm.Checked = True Then
                    TextBox_SCALE_PS.Text = (CDbl(TextBox_printScale.Text) / 1000) / viewport_scale
                Else
                    TextBox_SCALE_PS.Text = CDbl(TextBox_printScale.Text) / viewport_scale
                End If

            End If


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub TextBox_SCALE_PS_Click(sender As Object, e As System.EventArgs) Handles TextBox_SCALE_PS.Click
        Try
            If IsNumeric(TextBox_SCALE_PS.Text) = True And IsNumeric(TextBox_printScale.Text) = True Then
                Dim ps_scale As Double = CDbl(TextBox_SCALE_PS.Text)
                If RadioButton_mm.Checked = True Then
                    TextBox_viewport_Scale.Text = (CDbl(TextBox_printScale.Text) / 1000) / ps_scale
                Else
                    TextBox_viewport_Scale.Text = CDbl(TextBox_printScale.Text) / ps_scale
                End If

            End If


        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_transfer_info_Click(sender As Object, e As EventArgs) Handles Button_transfer_info.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension source:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        ascunde_butoanele_pentru_forms(Me, Colectie1)

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction




                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim Dim1 As Dimension = Ent1

                                Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                                Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                                Object_Prompt2.MessageForAdding = vbLf & "Select dimension destination:"
                                Object_Prompt2.SingleOnly = False
                                Rezultat2 = Editor1.GetSelection(Object_Prompt2)
                                If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    For i = 0 To Rezultat2.Value.Count - 1
                                        Dim Ent2 As Entity
                                        Ent2 = Rezultat2.Value.Item(i).ObjectId.GetObject(OpenMode.ForRead)
                                        If TypeOf Ent2 Is Dimension Then
                                            Dim Dim2 As Dimension = Ent2
                                            Dim2.UpgradeOpen()
                                            Dim2.DimensionText = Dim1.DimensionText
                                        End If


                                    Next
                                End If



                            End If



                            Trans1.Commit()

                        End Using

                    Else

                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_hide_show_ext_line_1_Click(sender As Object, e As EventArgs) Handles Button_hide_show_ext_line_1.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        ascunde_butoanele_pentru_forms(Me, Colectie1)

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction




                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim Dim1 As Dimension = Ent1
                                Dim1.UpgradeOpen()
                                If Dim1.Dimse1 = False Then
                                    Dim1.Dimse1 = True
                                Else
                                    Dim1.Dimse1 = False
                                End If
                            End If



                            Trans1.Commit()

                        End Using

                    Else

                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_hide_show_ext_line_2_Click(sender As Object, e As EventArgs) Handles Button_hide_show_ext_line_2.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        ascunde_butoanele_pentru_forms(Me, Colectie1)

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction




                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim Dim1 As Dimension = Ent1
                                Dim1.UpgradeOpen()
                                If Dim1.Dimse2 = False Then
                                    Dim1.Dimse2 = True
                                Else
                                    Dim1.Dimse2 = False
                                End If
                            End If



                            Trans1.Commit()

                        End Using

                    Else

                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_align_to_linear_Click(sender As Object, e As EventArgs) Handles Button_align_to_linear.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        ascunde_butoanele_pentru_forms(Me, Colectie1)

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction




                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is AlignedDimension Then
                                Dim Dim1 As AlignedDimension = Ent1
                                Dim Dim2 As New RotatedDimension
                                Dim2.AlternatePrefix = Dim1.AlternatePrefix
                                Dim2.AlternateSuffix = Dim1.AlternateSuffix
                                Dim2.Annotative = Dim1.Annotative
                                Dim2.CastShadows = Dim1.CastShadows
                                Dim2.Color = Dim1.Color
                                Dim2.ColorIndex = Dim1.ColorIndex
                                Dim2.Dimadec = Dim1.Dimadec
                                Dim2.Dimalt = Dim1.Dimalt
                                Dim2.Dimaltd = Dim1.Dimaltd
                                Dim2.Dimaltf = Dim1.Dimaltf
                                Dim2.Dimaltmzf = Dim1.Dimaltmzf
                                Dim2.Dimaltmzs = Dim1.Dimaltmzs
                                Dim2.Dimaltrnd = Dim1.Dimaltrnd
                                Dim2.Dimalttd = Dim1.Dimalttd
                                Dim2.Dimalttz = Dim1.Dimalttz
                                Dim2.Dimaltu = Dim1.Dimaltu
                                Dim2.Dimaltz = Dim1.Dimaltz
                                Dim2.Dimapost = Dim1.Dimapost
                                Dim2.Dimarcsym = Dim1.Dimarcsym
                                Dim2.Dimasz = Dim1.Dimasz
                                Dim2.Dimatfit = Dim1.Dimatfit
                                Dim2.Dimaunit = Dim1.Dimaunit
                                Dim2.Dimazin = Dim1.Dimazin
                                Dim2.Dimblk = Dim1.Dimblk
                                Dim2.Dimblk1 = Dim1.Dimblk1
                                Dim2.Dimblk2 = Dim1.Dimblk2
                                Dim2.DimBlockId = Dim1.DimBlockId
                                Dim2.Dimcen = Dim1.Dimcen
                                Dim2.Dimclrd = Dim1.Dimclrd
                                Dim2.Dimclre = Dim1.Dimclre
                                Dim2.Dimclrt = Dim1.Dimclrt
                                Dim2.Dimdec = Dim1.Dimdec
                                Dim2.Dimdle = Dim1.Dimdle
                                Dim2.Dimdli = Dim1.Dimdli
                                Dim2.Dimdsep = Dim1.Dimdsep
                                Dim2.DimensionStyle = Dim1.DimensionStyle
                                Dim2.DimensionText = Dim1.DimensionText
                                Dim2.Dimexe = Dim1.Dimexe
                                Dim2.Dimexo = Dim1.Dimexo
                                Dim2.Dimfrac = Dim1.Dimfrac
                                Dim2.Dimfxlen = Dim1.Dimfxlen
                                Dim2.DimfxlenOn = Dim1.DimfxlenOn
                                Dim2.Dimgap = Dim1.Dimgap
                                Dim2.Dimjogang = Dim1.Dimjogang
                                Dim2.Dimjust = Dim1.Dimjust
                                Dim2.Dimldrblk = Dim1.Dimldrblk
                                Dim2.Dimlfac = Dim1.Dimlfac
                                Dim2.Dimlim = Dim1.Dimlim
                                Dim2.DimLinePoint = Dim1.DimLinePoint
                                Dim2.Dimltex1 = Dim1.Dimltex1
                                Dim2.Dimltex2 = Dim1.Dimltex2
                                Dim2.Dimltype = Dim1.Dimltype
                                Dim2.Dimlunit = Dim1.Dimlunit
                                Dim2.Dimlwd = Dim1.Dimlwd
                                Dim2.Dimlwe = Dim1.Dimlwe
                                Dim2.Dimmzf = Dim1.Dimmzf
                                Dim2.Dimmzs = Dim1.Dimmzs
                                Dim2.Dimpost = Dim1.Dimpost
                                Dim2.Dimrnd = Dim1.Dimrnd
                                Dim2.Dimsah = Dim1.Dimsah
                                Dim2.Dimscale = Dim1.Dimscale
                                Dim2.Dimsd1 = Dim1.Dimsd1
                                Dim2.Dimsd2 = Dim1.Dimsd2
                                Dim2.Dimse1 = Dim1.Dimse1
                                Dim2.Dimse2 = Dim1.Dimse2
                                Dim2.Dimsoxd = Dim1.Dimsoxd
                                Dim2.Dimtad = Dim1.Dimtad
                                Dim2.Dimtdec = Dim1.Dimtdec
                                Dim2.Dimtfac = Dim1.Dimtfac
                                Dim2.Dimtfill = Dim1.Dimtfill
                                Dim2.Dimtfillclr = Dim1.Dimtfillclr
                                Dim2.Dimtih = Dim1.Dimtih
                                Dim2.Dimtix = Dim1.Dimtix
                                Dim2.Dimtm = Dim1.Dimtm
                                Dim2.Dimtmove = Dim1.Dimtmove
                                Dim2.Dimtofl = Dim1.Dimtofl
                                Dim2.Dimtoh = Dim1.Dimtoh
                                Dim2.Dimtol = Dim1.Dimtol
                                Dim2.Dimtolj = Dim1.Dimtolj
                                Dim2.Dimtp = Dim1.Dimtp
                                Dim2.Dimtsz = Dim1.Dimtsz
                                Dim2.Dimtvp = Dim1.Dimtvp
                                Dim2.Dimtxt = Dim1.Dimtxt
                                Dim2.Dimtxtdirection = Dim1.Dimtxtdirection
                                Dim2.Dimtzin = Dim1.Dimtzin
                                Dim2.Dimupt = Dim1.Dimupt
                                Dim2.Dimzin = Dim1.Dimzin
                                Dim2.DrawStream = Dim1.DrawStream
                                Dim2.DynamicDimension = Dim1.DynamicDimension
                                Dim2.EdgeStyleId = Dim1.EdgeStyleId
                                Dim2.Elevation = Dim1.Elevation
                                Dim2.FaceStyleId = Dim1.FaceStyleId
                                Dim2.ForceAnnoAllVisible = Dim1.ForceAnnoAllVisible
                                Dim2.HasSaveVersionOverride = Dim1.HasSaveVersionOverride
                                Dim2.HorizontalRotation = Dim1.HorizontalRotation
                                Dim2.Layer = Dim1.Layer
                                Dim2.Linetype = Dim1.Linetype
                                Dim2.LinetypeId = Dim1.LinetypeId
                                Dim2.LinetypeScale = Dim1.LinetypeScale
                                Dim2.LineWeight = Dim1.LineWeight
                                Dim2.Normal = Dim1.Normal
                                Dim2.Oblique = Dim1.Oblique
                                Dim2.OwnerId = Dim1.OwnerId
                                Dim2.Prefix = Dim1.Prefix
                                Dim2.Suffix = Dim1.Suffix
                                Dim2.TextAttachment = Dim1.TextAttachment
                                Dim2.TextLineSpacingFactor = Dim1.TextLineSpacingFactor
                                Dim2.TextLineSpacingStyle = Dim1.TextLineSpacingStyle
                                Dim2.TextPosition = Dim1.TextPosition
                                Dim2.TextRotation = Dim1.TextRotation
                                Dim2.TextStyleId = Dim1.TextStyleId
                                Dim2.Transparency = Dim1.Transparency
                                Dim2.UsingDefaultTextPosition = Dim1.UsingDefaultTextPosition
                                Dim2.Visible = Dim1.Visible
                                Dim2.VisualStyleId = Dim1.VisualStyleId
                                Dim2.XData = Dim1.XData
                                Dim2.XLine1Point = Dim1.XLine1Point
                                Dim2.XLine2Point = Dim1.XLine2Point
                                Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                                BTrecord.AppendEntity(Dim2)
                                Trans1.AddNewlyCreatedDBObject(Dim2, True)
                                Dim1.UpgradeOpen()
                                Dim1.Erase()
                            End If



                            Trans1.Commit()

                        End Using

                    Else

                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_double_arrow1_Click(sender As Object, e As EventArgs) Handles Button_double_arrow1.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        ascunde_butoanele_pentru_forms(Me, Colectie1)



                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim Dim1 As Dimension = Ent1
                                Dim1.UpgradeOpen()
                                Insereaza_block_table_record_in_drawing("dbl_open30.dwg", "open30_double")
                                Dim DBLarrID As ObjectId = Get_Arrow_dimension_ID("DIMBLK1", "open30_double")
                                Dim1.Dimsah = True
                                Dim1.Dimblk1 = DBLarrID

                            End If



                            Trans1.Commit()

                        End Using

                    Else
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_double_arrow2_Click(sender As Object, e As EventArgs) Handles Button_double_arrow2.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then

                        ascunde_butoanele_pentru_forms(Me, Colectie1)



                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim Dim1 As Dimension = Ent1
                                Dim1.UpgradeOpen()
                                Insereaza_block_table_record_in_drawing("dbl_open30.dwg", "open30_double")
                                Dim DBLarrID As ObjectId = Get_Arrow_dimension_ID("DIMBLK1", "open30_double")
                                Dim1.Dimsah = True
                                Dim1.Dimblk2 = DBLarrID

                            End If



                            Trans1.Commit()

                        End Using





                    Else
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If
                End If
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)

                    ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_open30_1_Click(sender As Object, e As EventArgs) Handles Button_open30_1.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then


                        ascunde_butoanele_pentru_forms(Me, Colectie1)



                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim Dim1 As Dimension = Ent1
                                Dim1.UpgradeOpen()

                                Dim DBLarrID As ObjectId = Get_Arrow_dimension_ID("DIMBLK1", "_Open30")
                                Dim1.Dimsah = True
                                Dim1.Dimblk1 = DBLarrID

                            End If



                            Trans1.Commit()

                        End Using




                    Else

                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button_open30_2_Click(sender As Object, e As EventArgs) Handles Button_open30_2.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument


                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                Editor1 = ThisDrawing.Editor

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select dimension object:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    If IsNothing(Rezultat1) = False Then


                        ascunde_butoanele_pentru_forms(Me, Colectie1)



                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(0)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Dimension Then
                                Dim Dim1 As Dimension = Ent1
                                Dim1.UpgradeOpen()

                                Dim DBLarrID As ObjectId = Get_Arrow_dimension_ID("DIMBLK2", "_Open30")
                                Dim1.Dimsah = True
                                Dim1.Dimblk2 = DBLarrID

                            End If



                            Trans1.Commit()

                        End Using




                    Else

                        Exit Sub
                    End If
                End If
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Dim Empty_array() As ObjectId
                Editor1.SetImpliedSelection(Empty_array)

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using ' asta e de la lock
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub Button_dim_rotated_Click(sender As Object, e As EventArgs) Handles Button_dim_rotated.Click
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify first point:")
            PP1.AllowNone = False
            Point1 = Editor1.GetPoint(PP1)

            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If

            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify second point:")
            PP2.UseBasePoint = True
            PP2.BasePoint = Point1.Value

            PP2.AllowNone = False
            Point2 = Editor1.GetPoint(PP2)

            If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                Exit Sub
            End If


            Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Dim Dimension1 As New RotatedDimension
                    Dim Rotatie As Double = GET_Bearing_rad(Point1.Value.TransformBy(CurentUCSmatrix).X, Point1.Value.TransformBy(CurentUCSmatrix).Y, Point2.Value.TransformBy(CurentUCSmatrix).X, Point2.Value.TransformBy(CurentUCSmatrix).Y)

                    Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim Jig1 As jig_dim_Class
                    Jig1 = New jig_dim_Class(New AlignedDimension, Point1.Value.TransformBy(CurentUCSmatrix), Point2.Value.TransformBy(CurentUCSmatrix))
                    Point3 = Jig1.BeginJig

                    If Point3 Is Nothing Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If


                    Dimension1.XLine1Point = Point1.Value.TransformBy(CurentUCSmatrix)
                    Dimension1.XLine2Point = Point2.Value.TransformBy(CurentUCSmatrix)
                    Dimension1.Rotation = Rotatie
                    Dimension1.DimLinePoint = Point3.Value

                    BTrecord.AppendEntity(Dimension1)
                    Trans1.AddNewlyCreatedDBObject(Dimension1, True)
                    Trans1.Commit()
                End Using
            End Using

            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub



End Class