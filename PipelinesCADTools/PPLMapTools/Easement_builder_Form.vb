Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Easement_builder_Form
    Dim Colectie1 As New Specialized.StringCollection
    

    Private Sub Button_offset_Click(sender As Object, e As EventArgs) Handles Button_offset.Click
        Dim Left1 As Double = CDbl(TextBox_left.Text)
        Dim Right1 As Double = CDbl(TextBox_right.Text)

        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        ascunde_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Dim RezultatCL As Autodesk.AutoCAD.EditorInput.PromptEntityResult
                Dim Object_PromptCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select centerline:")
                RezultatCL = Editor1.GetEntity(Object_PromptCL)
                If RezultatCL.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                    Exit Sub
                End If

                Dim PolyCL As Polyline
                If RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(RezultatCL) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim Ent1 As Entity
                            Ent1 = Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                PolyCL = Trans1.GetObject(RezultatCL.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Else
                                Editor1.WriteMessage(vbLf & "No Polyline")
                                Editor1.WriteMessage(vbLf & "Command:")
                                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                Exit Sub
                            End If
                        End Using
                    End If
                End If

                Dim Rezultat_parcel As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt_parcel As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt_parcel.MessageForAdding = vbLf & "Select parcel:"

                Object_Prompt_parcel.SingleOnly = True
                Rezultat_parcel = Editor1.GetSelection(Object_Prompt_parcel)

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    If IsNothing(PolyCL) = False And Rezultat_parcel.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        For i = 0 To Rezultat_parcel.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat_parcel.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline Then
                                Dim poly_parcel As Polyline = Ent1

                                Dim Object_colection1 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = PolyCL.GetOffsetCurves(-Left1)

                                Dim Poly_Left As New Polyline
                                Poly_Left = Object_colection1(0)


                                Dim Object_colection2 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = PolyCL.GetOffsetCurves(Right1)
                                Dim Poly_Right As New Polyline
                                Poly_Right = Object_colection2(0)



                                Dim Col_int_left As New Point3dCollection
                                poly_parcel.IntersectWith(Poly_Left, Intersect.OnBothOperands, Col_int_left, IntPtr.Zero, IntPtr.Zero)
                                Dim New_poly_left As New Polyline
                                Dim New_poly_right As New Polyline


                                If Col_int_left.Count > 1 Then

                                    Dim Point1 As New Point3d
                                    Point1 = Col_int_left(0)
                                    Dim Point2 As New Point3d
                                    Point2 = Col_int_left(Col_int_left.Count - 1)

                                    Dim Param1 As Double = Poly_Left.GetParameterAtPoint(Point1)
                                    Dim Param2 As Double = Poly_Left.GetParameterAtPoint(Point2)
                                    If Param1 > Param2 Then
                                        Dim T As Double = Param1
                                        Param1 = Param2
                                        Param2 = T
                                        Dim Pt As New Point3d
                                        Pt = Point1
                                        Point1 = Point2
                                        Point2 = Pt
                                    End If


                                    New_poly_left.AddVertexAt(0, New Point2d(Point1.X, Point1.Y), 0, 0, 0)

                                    Dim Index_left As Integer = 1
                                    For j = 0 To Poly_Left.NumberOfVertices - 1
                                        If j > Param1 And j < Param2 Then
                                            New_poly_left.AddVertexAt(Index_left, Poly_Left.GetPoint2dAt(j), 0, 0, 0)
                                            Index_left = Index_left + 1
                                        End If

                                    Next
                                    New_poly_left.AddVertexAt(Index_left, New Point2d(Point2.X, Point2.Y), 0, 0, 0)
                                    New_poly_left.ColorIndex = 40

                                    BTrecord.AppendEntity(New_poly_left)
                                    Trans1.AddNewlyCreatedDBObject(New_poly_left, True)




                                End If

                                Dim Col_int_right As New Point3dCollection
                                poly_parcel.IntersectWith(Poly_Right, Intersect.OnBothOperands, Col_int_right, IntPtr.Zero, IntPtr.Zero)
                                If Col_int_right.Count > 1 Then
                                    If Col_int_right.Count = 2 Then
                                        Dim Point1 As New Point3d
                                        Point1 = Col_int_right(0)
                                        Dim Point2 As New Point3d
                                        Point2 = Col_int_right(Col_int_right.Count - 1)

                                        Dim Param1 As Double = Poly_Right.GetParameterAtPoint(Point1)
                                        Dim Param2 As Double = Poly_Right.GetParameterAtPoint(Point2)
                                        If Param1 > Param2 Then
                                            Dim T As Double = Param1
                                            Param1 = Param2
                                            Param2 = T
                                            Dim Pt As New Point3d
                                            Pt = Point1
                                            Point1 = Point2
                                            Point2 = Pt
                                        End If

                                        New_poly_right.AddVertexAt(0, New Point2d(Point1.X, Point1.Y), 0, 0, 0)

                                        Dim Index_right As Integer = 1
                                        For j = 0 To Poly_Right.NumberOfVertices - 1
                                            If j > Param1 And j < Param2 Then
                                                New_poly_right.AddVertexAt(Index_right, Poly_Right.GetPoint2dAt(j), 0, 0, 0)
                                                Index_right = Index_right + 1
                                            End If

                                        Next
                                        New_poly_right.AddVertexAt(Index_right, New Point2d(Point2.X, Point2.Y), 0, 0, 0)
                                        New_poly_right.ColorIndex = 40

                                        BTrecord.AppendEntity(New_poly_right)
                                        Trans1.AddNewlyCreatedDBObject(New_poly_right, True)

                                    End If
                                End If

                                Dim Easement1 As New Polyline
                                Dim index_easement As Integer = 0
                                If IsNothing(New_poly_left) = False And IsNothing(New_poly_right) = False Then
                                    'Dim Param1 As Double = poly_parcel.GetParameterAtPoint(New_poly_left.GetPoint3dAt(0))
                                    'Dim Param2 As Double = poly_parcel.GetParameterAtPoint(New_poly_left.GetPoint3dAt(New_poly_left.NumberOfVertices - 1))
                                    'Dim Param3 As Double = poly_parcel.GetParameterAtPoint(New_poly_right.GetPoint3dAt(0))
                                    'Dim Param4 As Double = poly_parcel.GetParameterAtPoint(New_poly_right.GetPoint3dAt(New_poly_right.NumberOfVertices - 1))
                                End If


                            End If

                        Next

                        Trans1.Commit()
                    End If
                End Using


            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub
End Class