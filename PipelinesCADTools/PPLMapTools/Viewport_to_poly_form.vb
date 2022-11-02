Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Viewport_to_poly_form

    Dim Len1 As Double
    Dim Height1 As Double
    Dim Scale1 As Double
    Dim Angle1 As Double

    Dim vp_center_ms As New Point3d

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button_pick.Click
        Try

            '** added to command list
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Editor1 = ThisDrawing.Editor
            Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem
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
                            Angle1 = -Viewport1.TwistAngle
                            Len1 = Viewport1.Width
                            Height1 = Viewport1.Height
                            Scale1 = Viewport1.CustomScale
                            vp_center_ms = Application.GetSystemVariable("VIEWCTR") 'View0.Target
                            Button_draw.Visible = True
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager(0)
                            Exit For
                        End If
                    Next







                    Trans1.Commit()
                End Using ' asta e de la trans1
            End Using


        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
        Button_draw_Click(sender, e)
    End Sub

    Private Sub Viewport_to_poly_form_Click(sender As Object, e As EventArgs) Handles Me.Click
        Button_draw.Visible = False
        Len1 = 0
        Height1 = 0
        Scale1 = 0
        Angle1 = 0
    End Sub

    Private Sub Viewport_to_poly_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Button_draw.Visible = False
    End Sub

    Private Sub Button_draw_Click(sender As Object, e As EventArgs) Handles Button_draw.Click
        If Not Len1 = 0 And Not Height1 = 0 And Not Scale1 = 0 And Not Angle1 = 0 Then
            Try
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
                Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
                ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Editor1 = ThisDrawing.Editor


                Using LOCK2 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans2 As Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim BlockTable2 As BlockTable = Trans2.GetObject(ThisDrawing.Database.BlockTableId, OpenMode.ForRead)
                        Dim BTrecord2 As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans2.GetObject(BlockTable2(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Dim Poly1 As New Polyline
                        Creaza_layer_cu_linetype_si_lineweight("VW1", 1, "Continuous", LineWeight.ByLineWeightDefault, "", False, False)
                        Poly1.AddVertexAt(0, New Point2d(vp_center_ms.X - (Len1 / 2) / Scale1, vp_center_ms.Y - (Height1 / 2) / Scale1), 0, 1, 1)
                        Poly1.AddVertexAt(1, New Point2d(vp_center_ms.X - (Len1 / 2) / Scale1, vp_center_ms.Y + (Height1 / 2) / Scale1), 0, 1, 1)
                        Poly1.AddVertexAt(2, New Point2d(vp_center_ms.X + (Len1 / 2) / Scale1, vp_center_ms.Y + (Height1 / 2) / Scale1), 0, 1, 1)
                        Poly1.AddVertexAt(3, New Point2d(vp_center_ms.X + (Len1 / 2) / Scale1, vp_center_ms.Y - (Height1 / 2) / Scale1), 0, 1, 1)
                        Poly1.Layer = "VW1"
                        Poly1.Closed = True
                        Poly1.TransformBy(Matrix3d.Rotation(Angle1, Vector3d.ZAxis, vp_center_ms))
                        BTrecord2.AppendEntity(Poly1)
                        Trans2.AddNewlyCreatedDBObject(Poly1, True)
                        Trans2.Commit()
                    End Using
                End Using




                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Editor1.WriteMessage(vbLf & "Command:")

                If CheckBox_close.Checked = True Then
                    If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Count > 1 Then
                        DocumentExtension.CloseAndDiscard(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager(1))
                    End If

                End If

            Catch ex As Exception

                MsgBox(ex.Message)
            End Try


        End If
        Button_draw.Visible = False

        Len1 = 0
        Height1 = 0
        Scale1 = 0
        Angle1 = 0

    End Sub
End Class