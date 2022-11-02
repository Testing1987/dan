Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Protection_fence_form
    Dim Colectie1 As New Specialized.StringCollection
    Private Sub Button_draw_Click(sender As Object, e As EventArgs) Handles Button_draw.Click
        If IsNumeric(TextBox_buffer_dist.Text) = False Or _
                IsNumeric(TextBox_offset.Text) = False Or _
                IsNumeric(TextBox_max_dist.Text) = False Or _
                IsNumeric(TextBox_arrow_size.Text) = False Or _
                IsNumeric(TextBox_text_size.Text) = False Then


            MsgBox("Non numerical value")
            Exit Sub
        End If


        Dim Distanta As Double = CDbl(TextBox_buffer_dist.Text)
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

                Dim Rezultat_str As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt_str As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt_str.MessageForAdding = vbLf & "Select structures:"

                Object_Prompt_str.SingleOnly = False
                Rezultat_str = Editor1.GetSelection(Object_Prompt_str)

                Dim Rezultat_ws As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt_ws As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt_ws.MessageForAdding = vbLf & "Select existing workspace:"

                Object_Prompt_ws.SingleOnly = False
                Rezultat_ws = Editor1.GetSelection(Object_Prompt_ws)

                Dim colectie_linii_si_cercuri As New DBObjectCollection
                Dim colectie_poly_str As New DBObjectCollection
                Dim Poly_before_offset As New Polyline
                Dim Point_collection As New Point3dCollection
                Dim WS_collection As New DBObjectCollection
                Dim Continua1 As Boolean = False
                Dim Continua2 As Boolean = False

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    If IsNothing(PolyCL) = False And Rezultat_str.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat_ws.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        For i = 0 To Rezultat_str.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat_str.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline Then
                                Dim poly_str As Polyline = Ent1
                                colectie_poly_str.Add(Ent1)
                                For j = 0 To poly_str.NumberOfVertices - 1
                                    Dim Pt1 As New Point3d
                                    Pt1 = poly_str.GetPoint3dAt(j)
                                    If Point_collection.Contains(Pt1) = False Then
                                        Point_collection.Add(Pt1)
                                    End If
                                Next

                            End If
                        Next
                        For i = 0 To Rezultat_ws.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat_ws.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Polyline Then
                                WS_collection.Add(Ent1)
                            End If
                        Next
                        Continua1 = True
                        Trans1.Commit()
                    End If
                End Using
                Dim Pmax As New Point3d
                Dim Pmin As New Point3d
                If Continua1 = True Then
                    Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        If Point_collection.Count > 1 And WS_collection.Count > 0 Then
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                            Pmax = PolyCL.GetClosestPointTo(Point_collection(0), Vector3d.ZAxis, False)
                            Pmin = PolyCL.GetClosestPointTo(Point_collection(0), Vector3d.ZAxis, False)
                            Dim Chainage_max As Double = PolyCL.GetDistAtPoint(Pmax)
                            Dim Chainage_min As Double = PolyCL.GetDistAtPoint(Pmin)

                            For i = 1 To Point_collection.Count - 1
                                Dim P1 As New Point3d
                                P1 = PolyCL.GetClosestPointTo(Point_collection(i), Vector3d.ZAxis, False)
                                Dim Chainage1 As Double = PolyCL.GetDistAtPoint(P1)
                                If Chainage1 > Chainage_max Then
                                    Pmax = P1
                                    Chainage_max = Chainage1
                                End If
                                If Chainage1 < Chainage_min Then
                                    Pmin = P1
                                    Chainage_min = Chainage1
                                End If
                            Next
                            If Chainage_max <> Chainage_min Then
                                Dim Punct_min As New Point3d
                                Dim Punct_max As New Point3d

                                If Chainage_min - Distanta >= 0 And Chainage_max + Distanta <= PolyCL.Length Then
                                    Creaza_layer("NO PLOT", 40, "NO PLOT", False)

                                    Punct_min = PolyCL.GetPointAtDist(Chainage_min - Distanta)
                                    Dim Circle1 As New Circle(Punct_min, Vector3d.ZAxis, 10)

                                    Circle1.Layer = "NO PLOT"
                                    Circle1.LineWeight = LineWeight.LineWeight000
                                    Circle1.ColorIndex = 256

                                    BTrecord.AppendEntity(Circle1)
                                    Trans2.AddNewlyCreatedDBObject(Circle1, True)
                                    colectie_linii_si_cercuri.Add(Circle1)

                                    Punct_max = PolyCL.GetPointAtDist(Chainage_max + Distanta)
                                    Dim Circle2 As New Circle(Punct_max, Vector3d.ZAxis, 10)

                                    Circle2.Layer = "NO PLOT"
                                    Circle2.LineWeight = LineWeight.LineWeight000
                                    Circle2.ColorIndex = 256

                                    BTrecord.AppendEntity(Circle2)
                                    Trans2.AddNewlyCreatedDBObject(Circle2, True)
                                    colectie_linii_si_cercuri.Add(Circle2)

                                    For i = 0 To WS_collection.Count - 1
                                        Dim PolyWS As Polyline = WS_collection(i)
                                        Dim P_WS_MAX As New Point3d
                                        P_WS_MAX = PolyWS.GetClosestPointTo(Punct_max, Vector3d.ZAxis, False)
                                        Dim P_WS_MIN As New Point3d
                                        P_WS_MIN = PolyWS.GetClosestPointTo(Punct_min, Vector3d.ZAxis, False)

                                        Dim Linie_max As New Line(Punct_max, P_WS_MAX)
                                        Dim Linie_min As New Line(Punct_min, P_WS_MIN)


                                        Dim UNGHIMAX As Double
                                        Dim UNGHIMIN As Double

                                        Dim Param_max As Double = PolyWS.GetParameterAtPoint(P_WS_MAX)

                                        Dim LinieWS_max As New Line(PolyWS.GetPointAtParameter(Floor(Param_max)), PolyWS.GetPointAtParameter(Ceiling(Param_max)))
                                        UNGHIMAX = Linie_max.StartPoint.GetVectorTo(Linie_max.EndPoint).GetAngleTo(LinieWS_max.StartPoint.GetVectorTo(LinieWS_max.EndPoint))

                                        If Abs(Round(UNGHIMAX * 180 / PI, 0)) = 90 Then
                                            Dim Linie_max1 As New Line
                                            Dim Linie_max2 As New Line


                                            Linie_max.TransformBy(Matrix3d.Scaling(100, Punct_max))
                                            Dim Colint_max As New Point3dCollection
                                            Linie_max.IntersectWith(PolyWS, Intersect.OnBothOperands, Colint_max, IntPtr.Zero, IntPtr.Zero)

                                            If Colint_max.Count > 1 Then
                                                For j = 0 To Colint_max.Count - 1
                                                    If IsNothing(Linie_max1) = True Then
                                                        Linie_max1 = New Line(Punct_max, Colint_max(j))

                                                    Else
                                                        If Not Linie_max1.Length = 0 Then
                                                            Linie_max2 = New Line(Punct_max, Colint_max(j))

                                                        Else
                                                            Linie_max1 = New Line(Punct_max, Colint_max(j))

                                                        End If


                                                    End If
                                                Next
                                                If IsNothing(Linie_max1) = False And IsNothing(Linie_max2) = False Then
                                                    If Linie_max1.Length > Linie_max2.Length Then
                                                        Linie_max1.Layer = "NO PLOT"

                                                        Linie_max1.LineWeight = LineWeight.LineWeight000

                                                        Linie_max1.ColorIndex = 256

                                                        BTrecord.AppendEntity(Linie_max1)
                                                        Trans2.AddNewlyCreatedDBObject(Linie_max1, True)
                                                        colectie_linii_si_cercuri.Add(Linie_max1)
                                                    Else

                                                        Linie_max2.Layer = "NO PLOT"

                                                        Linie_max2.LineWeight = LineWeight.LineWeight000

                                                        Linie_max2.ColorIndex = 256

                                                        BTrecord.AppendEntity(Linie_max2)

                                                        Trans2.AddNewlyCreatedDBObject(Linie_max2, True)
                                                        colectie_linii_si_cercuri.Add(Linie_max2)
                                                    End If

                                                End If
                                            Else
                                                Linie_max.TransformBy(Matrix3d.Scaling(0.01, Punct_max))
                                                Linie_max.Layer = "NO PLOT"

                                                Linie_max.LineWeight = LineWeight.LineWeight000

                                                Linie_max.ColorIndex = 256

                                                BTrecord.AppendEntity(Linie_max)
                                                Trans2.AddNewlyCreatedDBObject(Linie_max, True)
                                                colectie_linii_si_cercuri.Add(Linie_max)
                                            End If

                                        End If



                                        Dim Param_min As Double = PolyWS.GetParameterAtPoint(P_WS_MIN)

                                        Dim LinieWS_min As New Line(PolyWS.GetPointAtParameter(Floor(Param_min)), PolyWS.GetPointAtParameter(Ceiling(Param_min)))
                                        UNGHIMIN = Linie_min.StartPoint.GetVectorTo(Linie_min.EndPoint).GetAngleTo(LinieWS_min.StartPoint.GetVectorTo(LinieWS_min.EndPoint))

                                        If Abs(Round(UNGHIMIN * 180 / PI, 0)) = 90 Then
                                            Dim Linie_min1 As New Line
                                            Dim Linie_min2 As New Line


                                            Linie_min.TransformBy(Matrix3d.Scaling(100, Punct_min))
                                            Dim Colint_min As New Point3dCollection
                                            Linie_min.IntersectWith(PolyWS, Intersect.OnBothOperands, Colint_min, IntPtr.Zero, IntPtr.Zero)

                                            If Colint_min.Count > 1 Then
                                                For j = 0 To Colint_min.Count - 1
                                                    If IsNothing(Linie_min1) = True Then
                                                        Linie_min1 = New Line(Punct_min, Colint_min(j))

                                                    Else
                                                        If Not Linie_min1.Length = 0 Then
                                                            Linie_min2 = New Line(Punct_min, Colint_min(j))

                                                        Else
                                                            Linie_min1 = New Line(Punct_min, Colint_min(j))

                                                        End If


                                                    End If
                                                Next
                                                If IsNothing(Linie_min1) = False And IsNothing(Linie_min2) = False Then
                                                    If Linie_min1.Length > Linie_min2.Length Then
                                                        Linie_min1.Layer = "NO PLOT"

                                                        Linie_min1.LineWeight = LineWeight.LineWeight000

                                                        Linie_min1.ColorIndex = 256

                                                        BTrecord.AppendEntity(Linie_min1)
                                                        Trans2.AddNewlyCreatedDBObject(Linie_min1, True)
                                                        colectie_linii_si_cercuri.Add(Linie_min1)
                                                    Else

                                                        Linie_min2.Layer = "NO PLOT"

                                                        Linie_min2.LineWeight = LineWeight.LineWeight000

                                                        Linie_min2.ColorIndex = 256
                                                        BTrecord.AppendEntity(Linie_min2)
                                                        Trans2.AddNewlyCreatedDBObject(Linie_min2, True)
                                                        colectie_linii_si_cercuri.Add(Linie_min2)
                                                    End If

                                                End If
                                            Else
                                                Linie_min.TransformBy(Matrix3d.Scaling(0.01, Punct_min))

                                                Linie_min.Layer = "NO PLOT"

                                                Linie_min.LineWeight = LineWeight.LineWeight000

                                                Linie_min.ColorIndex = 256
                                                BTrecord.AppendEntity(Linie_min)
                                                Trans2.AddNewlyCreatedDBObject(Linie_min, True)
                                                colectie_linii_si_cercuri.Add(Linie_min)
                                            End If

                                        End If
                                    Next
                                    Trans2.TransactionManager.QueueForGraphicsFlush()
                                    Trans2.Commit()
                                    Continua2 = True
                                End If
                            End If
                        End If
                    End Using
                End If





                If Continua2 = True Then
                    Using Trans3 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans3.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                        Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                        Dim Colectie_pt As New Point3dCollection
                        Dim Ask_for_point As Boolean = True
                        Dim Counter_for_pt As Integer = 1

                        Dim NEW_OSnap, Old_OSnap As Integer
                        Old_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

                        NEW_OSnap = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.End + Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Intersection

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)
                        Do Until Ask_for_point = False
                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select " & Counter_for_pt & " point:")
                            PP1.AllowNone = True
                            If Counter_for_pt > 1 Then
                                PP1.UseBasePoint = True
                                PP1.BasePoint = Colectie_pt(Counter_for_pt - 2)
                            End If
                            Point1 = Editor1.GetPoint(PP1)
                            If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Ask_for_point = False
                            Else
                                Colectie_pt.Add(Point1.Value)
                                Counter_for_pt = Counter_for_pt + 1
                            End If
                        Loop
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)

                        If Colectie_pt.Count > 1 Then
                            For i = 0 To Colectie_pt.Count - 1
                                Poly_before_offset.AddVertexAt(i, New Point2d(Colectie_pt(i).X, Colectie_pt(i).Y), 0, 2, 2)
                            Next
                            BTrecord.AppendEntity(Poly_before_offset)
                            Trans3.AddNewlyCreatedDBObject(Poly_before_offset, True)

                            Dim punct_pt_dir_offset As New Point3d((Pmax.X + Pmin.X) / 2, (Pmax.Y + Pmin.Y) / 2, (Pmax.Z + Pmin.Z) / 2)
                            Dim Directie_pt_offset As Double = Directie_offset(Poly_before_offset.ObjectId, punct_pt_dir_offset)

                            Dim Object_colection1 As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection = Poly_before_offset.GetOffsetCurves(CDbl(TextBox_offset.Text) * Directie_pt_offset)

                            Dim Poly_after_offset As New Polyline

                            Poly_after_offset = Object_colection1(0)
                            Dim Este_in_afara As Boolean = False

                            For i = 0 To colectie_poly_str.Count - 1
                                Dim polystr As Polyline = Trans3.GetObject(colectie_poly_str(i).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                For j = 0 To polystr.NumberOfVertices - 1
                                    Dim pTTT1 As New Point3d
                                    pTTT1 = polystr.GetPoint3dAt(j)
                                    Dim Clossest_pt As New Point3d
                                    Clossest_pt = Poly_before_offset.GetClosestPointTo(pTTT1, Vector3d.ZAxis, False)
                                    Dim Len1 As Double = pTTT1.GetVectorTo(Clossest_pt).Length
                                    If Len1 > CDbl(TextBox_max_dist.Text) Then
                                        Este_in_afara = True
                                    Else
                                        Este_in_afara = False
                                        Exit For
                                    End If
                                Next
                                If Este_in_afara = True Then
                                    MsgBox("Structure is outside of " & TextBox_max_dist.Text)
                                End If
                            Next





                            If Este_in_afara = False Then
                                BTrecord.AppendEntity(Poly_after_offset)
                                Trans3.AddNewlyCreatedDBObject(Poly_after_offset, True)
                                Trans3.TransactionManager.QueueForGraphicsFlush()


                                Dim P1 As New Point3d
                                P1 = PolyCL.GetClosestPointTo(Poly_before_offset.StartPoint, Vector3d.ZAxis, False)
                                Dim P2 As New Point3d
                                P2 = PolyCL.GetClosestPointTo(Poly_before_offset.EndPoint, Vector3d.ZAxis, False)
                                Dim TextS As Double = CDbl(TextBox_text_size.Text)
                                Dim Arrows As Double = CDbl(TextBox_arrow_size.Text)


                                If CheckBox_sta.Checked = True Then
                                    Dim Chainage1 As String = Get_chainage_feet_from_double(PolyCL.GetDistAtPoint(P1), 0)
                                    Dim Mleader1 As New MLeader
                                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.StartPoint, "STA = " & Chainage1, TextS, Arrows, Arrows, 2 * TextS, 3 * TextS)
                                    Mleader1.Linetype = "CONTINUOUS"
                                    Mleader1.LineWeight = LineWeight.LineWeight000

                                    Dim Mleader2 As New MLeader
                                    Dim Chainage2 As String = Get_chainage_feet_from_double(PolyCL.GetDistAtPoint(P2), 0)
                                    Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.EndPoint, "STA = " & Chainage2, TextS, Arrows, Arrows, 2 * TextS, 3 * TextS)
                                    Mleader2.Linetype = "CONTINUOUS"
                                    Mleader2.LineWeight = LineWeight.LineWeight000
                                End If

                                If CheckBox_MP.Checked = True Then
                                    Dim Chainage1 As Double = PolyCL.GetDistAtPoint(P1)
                                    Dim Mleader1 As New MLeader
                                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.StartPoint, "MP = " & Get_String_Rounded(Chainage1 / 5280, 2), TextS, Arrows, Arrows, 2 * TextS, 5 * TextS)
                                    Mleader1.Linetype = "CONTINUOUS"
                                    Mleader1.LineWeight = LineWeight.LineWeight000
                                    Dim Mleader2 As New MLeader
                                    Dim Chainage2 As Double = PolyCL.GetDistAtPoint(P2)
                                    Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.EndPoint, "MP = " & Get_String_Rounded(Chainage2 / 5280, 2), TextS, Arrows, Arrows, 2 * TextS, 5 * TextS)
                                    Mleader2.Linetype = "CONTINUOUS"
                                    Mleader2.LineWeight = LineWeight.LineWeight000
                                End If

                                Dim Acmap As Autodesk.Gis.Map.Platform.AcMapMap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap
                                Dim Curent_system As String = Acmap.GetMapSRS()
                                If CheckBox_lat_long.Checked = True Then
                                    If String.IsNullOrEmpty(Curent_system) = True Then
                                        MsgBox("Please set your coordinate system")
                                        CheckBox_lat_long.Checked = False
                                    End If
                                End If


                                If CheckBox_lat_long.Checked = True Then
                                    Dim String_LL84 As String = "GEOGCS[" & Chr(34) & "LL84" & Chr(34) & ",DATUM[" & Chr(34) & "WGS84" & Chr(34) & ",SPHEROID[" & Chr(34) & "WGS84" & Chr(34) & ",6378137.000,298.25722293]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"

                                    Dim Coord_factory1 As New OSGeo.MapGuide.MgCoordinateSystemFactory
                                    Dim CoordSys1 As OSGeo.MapGuide.MgCoordinateSystem = Coord_factory1.Create(Curent_system)
                                    Dim CoordSys2 As OSGeo.MapGuide.MgCoordinateSystem = Coord_factory1.Create(String_LL84)
                                    Dim Transform1 As OSGeo.MapGuide.MgCoordinateSystemTransform = Coord_factory1.GetTransform(CoordSys1, CoordSys2)



                                    Dim x1 As Double = Poly_after_offset.StartPoint.X
                                    Dim y1 As Double = Poly_after_offset.StartPoint.Y
                                    Dim Coord1 As OSGeo.MapGuide.MgCoordinate = Transform1.Transform(x1, y1)
                                    Dim Lat1 As Double = Coord1.Y
                                    Dim Long1 As Double = Coord1.X
                                    Dim Mleader1 As New MLeader
                                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.StartPoint, "Lon = " & Round(Long1, 4) & vbCrLf & "Lat = " & Get_String_Rounded(Lat1, 4), TextS, Arrows, Arrows, 4 * TextS, TextS)
                                    Mleader1.Linetype = "CONTINUOUS"
                                    Mleader1.LineWeight = LineWeight.LineWeight000

                                    Dim x2 As Double = Poly_after_offset.EndPoint.X
                                    Dim y2 As Double = Poly_after_offset.EndPoint.Y
                                    Dim Coord2 As OSGeo.MapGuide.MgCoordinate = Transform1.Transform(x2, y2)
                                    Dim Lat2 As Double = Coord2.Y
                                    Dim Long2 As Double = Coord2.X
                                    Dim Mleader2 As New MLeader
                                    Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Poly_after_offset.EndPoint, "Lon = " & Round(Long2, 4) & vbCrLf & "Lat = " & Get_String_Rounded(Lat2, 4), TextS, Arrows, Arrows, 4 * TextS, TextS)
                                    Mleader2.Linetype = "CONTINUOUS"
                                    Mleader2.LineWeight = LineWeight.LineWeight000

                                End If



                            End If


                        End If

                        If colectie_linii_si_cercuri.Count > 0 Then
                            For i = 0 To colectie_linii_si_cercuri.Count - 1
                                Dim Dbobj1 As DBObject = Trans3.GetObject(colectie_linii_si_cercuri(i).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite)
                                Dbobj1.Erase()
                            Next
                        End If
                        If IsNothing(Poly_before_offset) = False Then
                            Poly_before_offset.Erase()
                        End If


                        Trans3.Commit()

                    End Using
                End If

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