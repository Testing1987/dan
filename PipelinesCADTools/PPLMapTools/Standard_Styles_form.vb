Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Standard_Styles_form
    Dim TextStyle1 As TextStyleTableRecord
    Dim DimStyle1 As DimStyleTableRecord
    Dim MleaderStyle1 As MLeaderStyle
    Dim Viewport_cen As Point3d
    Dim Viewport_width As Double
    Dim Viewport_height As Double
    Dim Viewport_scale As Double
    Dim Viewport_twist As Double
    Dim Viewport_position As Point3d

    Private Sub Button_LOAD_STYLES_Click(sender As Object, e As EventArgs) Handles Button_LOAD_STYLES.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    For Each Id1 As ObjectId In Text_style_table
                        Dim style1 As TextStyleTableRecord = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), TextStyleTableRecord)
                        If IsNothing(style1) = False Then
                            If TextBox_text_style.Text.ToUpper = style1.Name.ToUpper Then
                                TextStyle1 = style1
                                ComboBox_text_style.Items.Add(style1.Name)
                                ComboBox_text_style.SelectedIndex = 0
                                Exit For
                            End If

                        End If

                    Next
                    Dim Dim_style_table As Autodesk.AutoCAD.DatabaseServices.DimStyleTable = Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    For Each Id1 As ObjectId In Dim_style_table
                        Dim style2 As DimStyleTableRecord = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), DimStyleTableRecord)
                        If IsNothing(style2) = False Then
                            If TextBox_dim_style.Text.ToUpper = style2.Name.ToUpper Then
                                DimStyle1 = style2
                                ComboBox_Dim_style.Items.Add(style2.Name)
                                ComboBox_Dim_style.SelectedIndex = 0
                                Exit For
                            End If
                        End If
                    Next

                    Dim Mleader_style_table As DBDictionary = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    If Mleader_style_table.Contains(TextBox_Mleader_style.Text) = True Then
                        Dim ID1 As ObjectId = Mleader_style_table.GetAt(TextBox_Mleader_style.Text)
                        MleaderStyle1 = TryCast(Trans1.GetObject(ID1, OpenMode.ForRead), MLeaderStyle)
                        ComboBox_Mleader_style.Items.Add(MleaderStyle1.Name)
                        ComboBox_Mleader_style.SelectedIndex = 0

                    End If

                   
                End Using
            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TRANSFER_STYLES_Click(sender As Object, e As EventArgs) Handles Button_TRANSFER_STYLES.Click
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    If IsNothing(TextStyle1) = False Then
                        If Text_style_table.Has(TextStyle1.Name) = True Then
                            Dim style1 As TextStyleTableRecord = TryCast(Trans1.GetObject(Text_style_table(TextStyle1.Name), OpenMode.ForWrite), TextStyleTableRecord)
                            style1.FileName = TextStyle1.FileName
                            style1.TextSize = TextStyle1.TextSize
                            style1.ObliquingAngle = TextStyle1.ObliquingAngle
                            style1.XScale = TextStyle1.XScale
                        Else
                            Dim style1 As New TextStyleTableRecord
                            style1 = TextStyle1.Clone
                            Text_style_table.UpgradeOpen()
                            Text_style_table.Add(style1)
                            Trans1.AddNewlyCreatedDBObject(style1, True)

                        End If
                    End If

                    Dim Dim_style_table As Autodesk.AutoCAD.DatabaseServices.DimStyleTable = Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    If IsNothing(DimStyle1) = False Then
                        If Dim_style_table.Has(DimStyle1.Name) = True Then
                            Dim style1 As DimStyleTableRecord = TryCast(Trans1.GetObject(Dim_style_table(DimStyle1.Name), OpenMode.ForWrite), DimStyleTableRecord)
                            style1.Dimadec = DimStyle1.Dimadec
                            style1.Dimalt = DimStyle1.Dimalt
                            style1.Dimaltd = DimStyle1.Dimaltd
                            style1.Dimaltf = DimStyle1.Dimaltf
                            style1.Dimaltrnd = DimStyle1.Dimaltrnd
                            style1.Dimalttd = DimStyle1.Dimalttd
                            style1.Dimalttz = DimStyle1.Dimalttz
                            style1.Dimaltu = DimStyle1.Dimaltu
                            style1.Dimaltz = DimStyle1.Dimaltz
                            style1.Dimapost = DimStyle1.Dimapost
                            style1.Dimarcsym = DimStyle1.Dimarcsym
                            style1.Dimasz = DimStyle1.Dimasz
                            style1.Dimatfit = DimStyle1.Dimatfit
                            style1.Dimaunit = DimStyle1.Dimaunit
                            style1.Dimazin = DimStyle1.Dimazin
                            style1.Dimcen = DimStyle1.Dimcen
                            style1.Dimclrd = DimStyle1.Dimclrd
                            style1.Dimclre = DimStyle1.Dimclre
                            style1.Dimclrt = DimStyle1.Dimclrt
                            style1.Dimdec = DimStyle1.Dimdec
                            style1.Dimdle = DimStyle1.Dimdle
                            style1.Dimdli = DimStyle1.Dimdli
                            style1.Dimdsep = DimStyle1.Dimdsep
                            style1.Dimexe = DimStyle1.Dimexe
                            style1.Dimexo = DimStyle1.Dimexo
                            style1.Dimfrac = DimStyle1.Dimfrac
                            style1.Dimfxlen = DimStyle1.Dimfxlen
                            style1.DimfxlenOn = DimStyle1.DimfxlenOn
                            style1.Dimgap = DimStyle1.Dimgap
                            style1.Dimjogang = DimStyle1.Dimjogang
                            style1.Dimjust = DimStyle1.Dimjust
                            style1.Dimblk = DimStyle1.Dimblk
                            style1.Dimldrblk = DimStyle1.Dimldrblk
                            style1.Dimlfac = DimStyle1.Dimlfac
                            style1.Dimlim = DimStyle1.Dimlim
                            style1.Dimlunit = DimStyle1.Dimlunit
                            style1.Dimlwd = DimStyle1.Dimlwd
                            style1.Dimlwe = DimStyle1.Dimlwe
                            style1.Dimpost = DimStyle1.Dimpost
                            style1.Dimrnd = DimStyle1.Dimrnd
                            style1.Dimsah = DimStyle1.Dimsah
                            style1.Dimscale = DimStyle1.Dimscale
                            style1.Dimsd1 = DimStyle1.Dimsd1
                            style1.Dimsd2 = DimStyle1.Dimsd2
                            style1.Dimse1 = DimStyle1.Dimse1
                            style1.Dimse2 = DimStyle1.Dimse2
                            style1.Dimsoxd = DimStyle1.Dimsoxd
                            style1.Dimtad = DimStyle1.Dimtad
                            style1.Dimtdec = DimStyle1.Dimtdec
                            style1.Dimtfac = DimStyle1.Dimtfac
                            style1.Dimtfill = DimStyle1.Dimtfill
                            style1.Dimtfillclr = DimStyle1.Dimtfillclr
                            style1.Dimtih = DimStyle1.Dimtih
                            style1.Dimtix = DimStyle1.Dimtix
                            style1.Dimtm = DimStyle1.Dimtm
                            style1.Dimtmove = DimStyle1.Dimtmove
                            style1.Dimtofl = DimStyle1.Dimtofl
                            style1.Dimtoh = DimStyle1.Dimtoh
                            style1.Dimtol = DimStyle1.Dimtol
                            style1.Dimtolj = DimStyle1.Dimtolj
                            style1.Dimtp = DimStyle1.Dimtp
                            style1.Dimtsz = DimStyle1.Dimtsz
                            style1.Dimtvp = DimStyle1.Dimtvp
                            style1.Dimtxsty = DimStyle1.Dimtxsty
                            style1.Dimtxt = DimStyle1.Dimtxt
                            style1.Dimtxtdirection = DimStyle1.Dimtxtdirection
                            style1.Dimtzin = DimStyle1.Dimtzin
                            style1.Dimupt = DimStyle1.Dimupt
                            style1.Dimzin = DimStyle1.Dimzin



                        Else
                            Dim style1 As New DimStyleTableRecord
                            style1 = DimStyle1.Clone
                            Dim_style_table.UpgradeOpen()
                            Dim_style_table.Add(style1)
                            Trans1.AddNewlyCreatedDBObject(style1, True)

                        End If

                        Dim Mleader_style_table As DBDictionary = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If IsNothing(MleaderStyle1) = False Then
                            If Mleader_style_table.Contains(TextBox_Mleader_style.Text) = True Then
                                Dim ID1 As ObjectId = Mleader_style_table.GetAt(TextBox_Mleader_style.Text)
                                Dim style1 As MLeaderStyle = TryCast(Trans1.GetObject(ID1, OpenMode.ForWrite), MLeaderStyle)
                                style1.Annotative = MleaderStyle1.Annotative
                                style1.ArrowSize = MleaderStyle1.ArrowSize
                                style1.BreakSize = MleaderStyle1.BreakSize
                                style1.DoglegLength = MleaderStyle1.DoglegLength
                                style1.LeaderLineColor = MleaderStyle1.LeaderLineColor
                                style1.TextColor = MleaderStyle1.TextColor
                                style1.TextHeight = MleaderStyle1.TextHeight
                                style1.TextStyleId = MleaderStyle1.TextStyleId
                                style1.ArrowSymbolId = MleaderStyle1.ArrowSymbolId
                                style1.BlockColor = MleaderStyle1.BlockColor
                                style1.BlockRotation = MleaderStyle1.BlockRotation
                                style1.BlockScale = MleaderStyle1.BlockScale
                                style1.ContentType = MleaderStyle1.ContentType
                                style1.DrawLeaderOrderType = MleaderStyle1.DrawLeaderOrderType
                                style1.DrawMLeaderOrderType = MleaderStyle1.DrawMLeaderOrderType
                                style1.EnableBlockRotation = MleaderStyle1.EnableBlockRotation
                                style1.EnableBlockScale = MleaderStyle1.EnableBlockScale
                                style1.EnableDogleg = MleaderStyle1.EnableDogleg
                                style1.EnableFrameText = MleaderStyle1.EnableFrameText
                                style1.EnableLanding = MleaderStyle1.EnableLanding
                                style1.ExtendLeaderToText = MleaderStyle1.ExtendLeaderToText
                                style1.TextAlignAlwaysLeft = MleaderStyle1.TextAlignAlwaysLeft
                                style1.LandingGap = MleaderStyle1.LandingGap
                                style1.LeaderLineType = MleaderStyle1.LeaderLineType
                                style1.LeaderLineWeight = MleaderStyle1.LeaderLineWeight
                                style1.MaxLeaderSegmentsPoints = MleaderStyle1.MaxLeaderSegmentsPoints
                                style1.Scale = 1 'MleaderStyle1.Scale
                                style1.TextAlignAlwaysLeft = MleaderStyle1.TextAlignAlwaysLeft
                                style1.TextAlignmentType = MleaderStyle1.TextAlignmentType
                                style1.TextAngleType = MleaderStyle1.TextAngleType
                                style1.TextAttachmentDirection = MleaderStyle1.TextAttachmentDirection
                                style1.TextAttachmentType = MleaderStyle1.TextAttachmentType

                            Else
                                Dim style1 As New MLeaderStyle
                                style1 = MleaderStyle1.Clone
                                style1.PostMLeaderStyleToDb(ThisDrawing.Database, MleaderStyle1.Name)
                                Trans1.AddNewlyCreatedDBObject(style1, True)


                            End If
                        End If


                    End If


                    Trans1.Commit()

                End Using
            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_read_viewport_Click(sender As Object, e As EventArgs) Handles Button_read_viewport.Click
        Try

            TRANSFER_STYLES_Click(sender, e)
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Editor1 = ThisDrawing.Editor

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()




            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Empty_array() As ObjectId
                    Editor1.SetImpliedSelection(Empty_array)



                    Dim Rezultat_viewport As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Prompt_viewport As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Prompt_viewport.MessageForAdding = vbLf & "Select the viewport"

                    Prompt_viewport.SingleOnly = False
                    Rezultat_viewport = Editor1.GetSelection(Prompt_viewport)


                    If Rezultat_viewport.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim BlockTable As BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                    Dim BTrecordPS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable(BlockTableRecord.PaperSpace), OpenMode.ForWrite)

                    Dim Dim_style_table As Autodesk.AutoCAD.DatabaseServices.DimStyleTable = Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim DimStyleId As ObjectId
                    If IsNothing(DimStyle1) = False Then
                        If Dim_style_table.Has(DimStyle1.Name) = True Then
                            DimStyleId = Dim_style_table(DimStyle1.Name)
                        End If
                    End If



                    Dim Mleader_style_table As DBDictionary = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim IDml As ObjectId

                    If Mleader_style_table.Contains(TextBox_Mleader_style.Text) = True Then
                        IDml = Mleader_style_table.GetAt(TextBox_Mleader_style.Text)
                    End If




                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim textStyleId As ObjectId
                    If IsNothing(TextStyle1) = False Then
                        If Text_style_table.Has(TextStyle1.Name) = True Then
                            textStyleId = Text_style_table(TextStyle1.Name)
                        End If
                    End If


                    For i = 0 To Rezultat_viewport.Value.Count - 1
                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat_viewport.Value.Item(i)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Viewport Then
                            Dim Viewport1 As Viewport = Ent1
                            Editor1.SwitchToModelSpace()
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CVPORT", Viewport1.Number)
                            Editor1.CurrentUserCoordinateSystem = WCS_align()
                            'Dim GraphicsManager As Autodesk.AutoCAD.GraphicsSystem.Manager = ThisDrawing.GraphicsManager
                            'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetGsView(CShort(Application.GetSystemVariable("CVPORT")), True) ' acad 2013
                            'Dim View0 As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.GetCurrentAcGsView(CShort(Application.GetSystemVariable("CVPORT"))) ' acad 2015
                            TextBox_rotation.Text = Round(Viewport1.TwistAngle, 3)
                            Viewport_twist = Viewport1.TwistAngle
                            Viewport_width = Viewport1.Width
                            Viewport_height = Viewport1.Height
                            TextBox_viewport_scale.Text = Round(Viewport1.CustomScale, 3)
                            Viewport_scale = Viewport1.CustomScale
                            Viewport_cen = Application.GetSystemVariable("VIEWCTR") 'View0.Target
                            Viewport_position = Viewport1.CenterPoint

                            Dim Rezultat_Xref As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Prompt_Xref As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Prompt_Xref.MessageForAdding = vbLf & "Select the Xref"

                            Prompt_Xref.SingleOnly = False
                            Rezultat_Xref = Editor1.GetSelection(Prompt_Xref)
                            If Rezultat_Xref.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.SetImpliedSelection(Empty_array)
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                            Dim XrefIdColl As New ObjectIdCollection
                            Dim Colectie_nume As New Specialized.StringCollection

                            For j = 0 To Rezultat_Xref.Value.Count - 1
                                Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj2 = Rezultat_Xref.Value.Item(j)
                                Dim Ent2 As Entity
                                Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                    Dim Block1 As BlockReference = Ent2
                                    If BlockTable.Has(Block1.Name) = True Then

                                        Dim BTR As BlockTableRecord = Trans1.GetObject(BlockTable(Block1.Name), OpenMode.ForRead)
                                        If BTR.IsResolved And BTR.IsFromExternalReference = True Then
                                            XrefIdColl.Add(BTR.ObjectId)
                                            Colectie_nume.Add(Block1.Name)
                                        End If
                                    End If


                                End If
                            Next

                            If XrefIdColl.Count > 0 Then
                                ThisDrawing.Database.BindXrefs(XrefIdColl, True)
                            End If


                            Dim Object_chspace_col As New DBObjectCollection

                            For Each Id1 As ObjectId In BTrecord
                                Dim Ent3 As Entity = Trans1.GetObject(Id1, OpenMode.ForRead)
                                If TypeOf Ent3 Is BlockReference Then
                                    Dim Block1 As BlockReference = Ent3
                                    For k = 0 To Colectie_nume.Count - 1
                                        If Block1.Name = Colectie_nume(k) Then
                                            Block1.UpgradeOpen()
                                            Dim DBCOL As New DBObjectCollection
                                            Block1.Explode(DBCOL)
                                            BTrecord.UpgradeOpen()
                                            For Each Ent4 As Entity In DBCOL
                                                BTrecord.AppendEntity(Ent4)
                                                Trans1.AddNewlyCreatedDBObject(Ent4, True)
                                                Object_chspace_col.Add(Ent4)
                                            Next


                                            Block1.Erase()

                                        End If
                                    Next

                                End If
                            Next

                            If IsNothing(Object_chspace_col) = False Then

                                For j = 0 To Object_chspace_col.Count - 1
                                    Dim Vector_translatie As New Vector3d
                                    Vector_translatie = Viewport_cen.GetVectorTo(Viewport_position)
                                    Vector_translatie.TransformBy(Matrix3d.Scaling(Viewport_scale, Viewport_cen))

                                    Dim ent5 As Entity = Object_chspace_col(j)
                                    Dim entPS As Entity = ent5.Clone
                                    entPS.TransformBy(Matrix3d.Scaling(Viewport_scale, Viewport_cen))
                                    entPS.TransformBy(Matrix3d.Rotation(-Viewport_twist, Vector3d.ZAxis, Viewport_cen))
                                    entPS.TransformBy(Matrix3d.Displacement(Vector_translatie))
                                    If TypeOf ent5 Is Dimension Then
                                        Dim Dim1 As Dimension = entPS
                                        Dim1.DimensionStyle = DimStyleId
                                        Dim1.TextStyleId = textStyleId
                                        Dim1.Dimlfac = 1 / Viewport_scale
                                    End If

                                    If TypeOf ent5 Is MLeader Then
                                        ent5.UpgradeOpen()
                                        Dim ML As MLeader = ent5
                                        ML.TextStyleId = textStyleId

                                    End If
                                    If TypeOf ent5 Is Leader Then
                                        Dim ML As Leader = entPS
                                        ML.DimensionStyle = DimStyleId

                                    End If
                                    If TypeOf ent5 Is DBText Then
                                        Dim txt As DBText = entPS
                                        txt.TextStyleId = textStyleId
                                    End If
                                    If TypeOf ent5 Is MText Then
                                        Dim mtxt As MText = entPS
                                        mtxt.TextStyleId = textStyleId
                                    End If

                                    If Not TypeOf ent5 Is MLeader Then
                                        BTrecordPS.AppendEntity(entPS)
                                        Trans1.AddNewlyCreatedDBObject(entPS, True)

                                        Dim Obiect_de_sters As DBObject
                                        Obiect_de_sters = Trans1.GetObject(Object_chspace_col(j).ObjectId, OpenMode.ForWrite)
                                        Obiect_de_sters.Erase()
                                    End If



                                Next


                            End If





                            Exit For
                        End If
                    Next







                    Trans1.Commit()
                End Using ' asta e de la trans1
            End Using


        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
    End Sub
End Class