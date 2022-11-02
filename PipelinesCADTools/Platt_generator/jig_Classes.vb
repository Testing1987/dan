Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Math

Public Class Jig_rectangle_viewport
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig

    Dim PunctM As New Point3d

    Dim Vw_width As Double
    Dim Vw_height As Double
    Dim Rotate_poly As Boolean

    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal VVw_width As Double, ByVal VVw_height As Double, ByVal Rotate As Boolean) As PromptPointResult
        Vw_width = VVw_width
        Vw_height = VVw_height
        Rotate_poly = Rotate
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptPointResult1 = ed.Drag(Me)

        Do
            Select Case PromptPointResult1.Status
                Case PromptStatus.OK
                    Return PromptPointResult1
                    Exit Do
            End Select
        Loop While PromptPointResult1.Status <> PromptStatus.Cancel
        Return PromptPointResult1
    End Function

    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus

        PromptPointResult1 = prompts.AcquirePoint(vbLf & "Specify target point for viewport:")


        If PromptPointResult1.Value.IsEqualTo(PunctM) Then
            Return SamplerStatus.NoChange
        Else
            PunctM = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
        Dim Vw_scale As Double = Platt_Generator_form.Vw_scale
        Dim Rotation1 As Double = Platt_Generator_form.Rotatie

        Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Poly1.AddVertexAt(0, New Point2d(PunctM.X - 0.5 * Vw_width / Vw_scale, PunctM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(1, New Point2d(PunctM.X + 0.5 * Vw_width / Vw_scale, PunctM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(2, New Point2d(PunctM.X + 0.5 * Vw_width / Vw_scale, PunctM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(3, New Point2d(PunctM.X - 0.5 * Vw_width / Vw_scale, PunctM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(4, New Point2d(PunctM.X - 0.5 * Vw_width / Vw_scale, PunctM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)

        If Rotate_poly = True Then
            Poly1.TransformBy(Matrix3d.Rotation(Rotation1, Vector3d.ZAxis, PunctM))
        End If


        draw.Geometry.Polyline(Poly1, 0, 4)


        Dim Poly2 As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Dim Poly3 As Polyline = Platt_Generator_form.Poly_parc_for_jig

        If IsNothing(Poly3) = False Then
            For i = 0 To Poly3.NumberOfVertices - 1
                Poly2.AddVertexAt(i, Poly3.GetPoint2dAt(i), 0, 10, 10)
            Next
            draw.Geometry.Polyline(Poly2, 0, Poly3.NumberOfVertices - 1)
        End If
        If Not Platt_Generator_form.pt_CERC = New Point3d() Then
            draw.Geometry.Circle(Platt_Generator_form.pt_CERC, 200, Vector3d.ZAxis)
        End If

    End Function
End Class


Public Class EntityRotateJigger
    Inherits EntityJig
    Dim Cancel_rot As Double = 0
#Region "Fields"

    Public mCurJigFactorIndex As Integer = 1

    Private mRotation As Double = 0.0
    ' Factor #1
    Private mBasePoint As New Point3d()
    Private mLastMatrix As New Matrix3d()
    Private mLastAngle As Double = 0.0
#End Region

#Region "Constructors"

    Public Sub New(ent As Entity, basePoint As Point3d)
        MyBase.New(ent)
        mBasePoint = basePoint
    End Sub

#End Region

#Region "Properties"

    Private ReadOnly Property Editor() As Editor
        Get
            Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        End Get
    End Property

    Private ReadOnly Property UCS() As Matrix3d
        Get
            Return Editor.CurrentUserCoordinateSystem
        End Get
    End Property

#End Region

#Region "Overrides"

    Protected Overrides Function Update() As Boolean

        Cancel_rot = Cancel_rot + mRotation - mLastAngle

        Dim basePt As New Point3d(mBasePoint.X, mBasePoint.Y, mBasePoint.Z)
        Dim mat As Matrix3d = Matrix3d.Rotation(mRotation - mLastAngle, Vector3d.ZAxis.TransformBy(UCS), basePt.TransformBy(UCS))
        Entity.TransformBy(mat)
        mLastAngle = mRotation



        Return True
    End Function

    Protected Overrides Function Sampler(prompts As JigPrompts) As SamplerStatus
        Select Case mCurJigFactorIndex
            Case 1
                Dim prOptions1 As New JigPromptAngleOptions(vbLf & "Rotation angle:")
                prOptions1.BasePoint = mBasePoint
                prOptions1.UseBasePoint = True
                prOptions1.UserInputControls = UserInputControls.NullResponseAccepted + UserInputControls.GovernedByOrthoMode



                Dim prResult1 As PromptDoubleResult = prompts.AcquireAngle(prOptions1)


                If prResult1.Status = PromptStatus.Cancel Then
                    Dim basePt As New Point3d(mBasePoint.X, mBasePoint.Y, mBasePoint.Z)
                    Dim mat As Matrix3d = Matrix3d.Rotation(-Cancel_rot, Vector3d.ZAxis.TransformBy(UCS), basePt.TransformBy(UCS))
                    Entity.TransformBy(mat)

                    Return SamplerStatus.Cancel
                End If
                If prResult1.Status = PromptStatus.Keyword Then

                End If
                If prResult1.Value.Equals(mRotation) Then
                    Return SamplerStatus.NoChange
                Else
                    mRotation = prResult1.Value
                    Return SamplerStatus.OK
                End If
            Case Else
                Exit Select
        End Select

        Return SamplerStatus.OK
    End Function

#End Region

#Region "Methods to Call"

    Public Shared Function Jig(ent As Entity, basePt As Point3d) As Boolean
        Dim jigger As EntityRotateJigger = Nothing
        Try
            jigger = New EntityRotateJigger(ent, basePt)
            Dim pr As PromptResult

            Do
                pr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.Drag(jigger)
                ' Add keyword handling code below

                If pr.Status = PromptStatus.Keyword Then
                Else

                    jigger.mCurJigFactorIndex += 1
                End If
            Loop While pr.Status <> PromptStatus.Cancel AndAlso pr.Status <> PromptStatus.[Error] AndAlso jigger.mCurJigFactorIndex <= 1

            If pr.Status = PromptStatus.Cancel OrElse pr.Status = PromptStatus.[Error] Then
                If jigger IsNot Nothing AndAlso jigger.Entity IsNot Nothing Then
                    jigger.Entity.Dispose()
                End If

                Return False
            Else

                Return True
            End If
        Catch
            If jigger IsNot Nothing AndAlso jigger.Entity IsNot Nothing Then
                jigger.Entity.Dispose()
            End If

            Return False
        End Try
    End Function

#End Region

End Class




Public Class Draw_JIG1
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig
    Dim Basept As Point3d
    Dim Base_start_point As Point3d
    Dim Inaltimea1 As Double


    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal Punct1 As Point3d, ByVal Scalefactor As Double) As PromptPointResult
        Base_start_point = Punct1
        Inaltimea1 = Scalefactor
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptPointResult1 = ed.Drag(Me)
        Do
            Select Case PromptPointResult1.Status
                Case PromptStatus.OK
                    Return PromptPointResult1
                    Exit Do
            End Select
        Loop While PromptPointResult1.Status <> PromptStatus.Cancel
        Return PromptPointResult1
    End Function

    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus

        PromptPointResult1 = prompts.AcquirePoint("Pick second point : ")

        If PromptPointResult1.Value.IsEqualTo(Basept) Then
            Return SamplerStatus.NoChange
        Else
            Basept = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean

        Dim x0 As Double = Base_start_point.X
        Dim y0 As Double = Base_start_point.Y

        Dim x As Double = Basept.X
        Dim y As Double = Basept.Y



        Dim Length1 As Double = (Inaltimea1) / 500
        Dim Wdth1 As Double = (Inaltimea1) / 1250
        Dim Dist1 As Double = ((x0 - x) ^ 2 + (y0 - y) ^ 2) ^ 0.5


        Dim xm, ym As Double
        Dim x1, y1 As Double

        Dim Bear1 As Double = GET_Bearing_rad(x0, y0, x, y)

        If Bear1 < PI / 2 Then
            x1 = Length1 * Cos(Bear1)
            y1 = Length1 * Sin(Bear1)
            xm = x0 + x1
            ym = y0 + y1
        End If

        If Bear1 = PI / 2 Then
            x1 = 0
            y1 = Length1
            xm = x0
            ym = y0 + y1
        End If

        If Bear1 > PI / 2 And Bear1 < PI Then
            x1 = Length1 * Cos(PI - Bear1)
            y1 = Length1 * Sin(PI - Bear1)
            xm = x0 - x1
            ym = y0 + y1
        End If

        If Bear1 = PI Then
            x1 = Length1
            y1 = 0
            xm = x0 - x1
            ym = y0
        End If

        If Bear1 > PI And Bear1 < 3 * PI / 2 Then
            x1 = Length1 * Cos(Bear1 - PI)
            y1 = Length1 * Sin(Bear1 - PI)
            xm = x0 - x1
            ym = y0 - y1
        End If

        If Bear1 = 3 * PI / 2 Then
            x1 = 0
            y1 = Length1
            xm = x0
            ym = y0 - y1
        End If

        If Bear1 > 3 * PI / 2 And Bear1 < 2 * PI Then
            x1 = Length1 * Cos(2 * PI - Bear1)
            y1 = Length1 * Sin(2 * PI - Bear1)
            xm = x0 + x1
            ym = y0 - y1
        End If

        If Bear1 = 2 * PI Or Bear1 = 0 Then
            x1 = Length1
            y1 = 0
            xm = x0 + x1
            ym = y0
        End If



        Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Poly1.AddVertexAt(0, New Point2d(x0, y0), 0, 0, Wdth1)
        Poly1.AddVertexAt(1, New Point2d(xm, ym), 0, 0, 0)
        Poly1.AddVertexAt(2, New Point2d(x, y), 0, 0, 0)

        draw.Geometry.Polyline(Poly1, 0, 2)

        'draw.Geometry.WorldLine(Basept1, Basept)


    End Function
End Class


Public Class Draw_JIG2
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig
    Dim Basept As Point3d
    Dim Base_start_point As Point3d
    Dim Base_mid_point As Point3d
    Dim Inaltimea1 As Double

    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal Punct1 As Point3d, ByVal Punct2 As Point3d, ByVal Scalefactor As Double) As PromptPointResult
        Base_start_point = Punct1
        Base_mid_point = Punct2
        Inaltimea1 = Scalefactor

        'Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", draw1.Existing_OSnap)

        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptPointResult1 = ed.Drag(Me)
        Do
            Select Case PromptPointResult1.Status
                Case PromptStatus.OK
                    Return PromptPointResult1
                    Exit Do
            End Select
        Loop While PromptPointResult1.Status <> PromptStatus.Cancel
        Return PromptPointResult1
    End Function

    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus

        PromptPointResult1 = prompts.AcquirePoint("Pick third point : ")

        If PromptPointResult1.Value.IsEqualTo(Basept) Then
            Return SamplerStatus.NoChange
        Else
            Basept = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean

        Dim x0 As Double = Base_start_point.X
        Dim y0 As Double = Base_start_point.Y

        Dim x2 As Double = Base_mid_point.X
        Dim y2 As Double = Base_mid_point.Y

        Dim x3 As Double = Basept.X
        Dim y3 As Double = Basept.Y


        Dim Length1 As Double = (Inaltimea1) / 500
        Dim Wdth1 As Double = (Inaltimea1) / 1250
        Dim Dist1 As Double = ((x0 - x2) ^ 2 + (y0 - y2) ^ 2) ^ 0.5

        Dim x1, y1 As Double
        x1 = x1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)
        y1 = y1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)

        Dim Bulge1 As Double = 0

        Bulge1 = Bulge_for_arc_leader(x0, y0, x2, y2, x3, y3, Inaltimea1)

        Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Poly1.AddVertexAt(0, New Point2d(x0, y0), 0, 0, Wdth1)
        Poly1.AddVertexAt(1, New Point2d(x1, y1), Bulge1, 0, 0)
        Poly1.AddVertexAt(2, New Point2d(x3, y3), 0, 0, 0)
        draw.Geometry.Polyline(Poly1, 0, 2)




        'draw.Geometry.WorldLine(Basept1, Basept)


    End Function
End Class

Public Class JIG_ACOLADE
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig
    Dim Point1 As Point3d
    Dim Point2 As Point3d
    Dim Point3 As Point3d
    Dim Arr_len As Double
    Dim Arr_width As Double
    Dim Rad_small As Double

    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal Punct1 As Point3d, ByVal Punct2 As Point3d, ByVal Arrow_length As Double, ByVal Arrow_width As Double, ByVal small_radius As Double) As PromptPointResult
        Point1 = Punct1
        Point2 = Punct2

        Arr_len = Arrow_length
        Arr_width = Arrow_width
        Rad_small = small_radius

        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptPointResult1 = ed.Drag(Me)
        Do
            Select Case PromptPointResult1.Status
                Case PromptStatus.OK
                    Return PromptPointResult1
                    Exit Do
            End Select
        Loop While PromptPointResult1.Status <> PromptStatus.Cancel
        Return PromptPointResult1
    End Function

    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus

        PromptPointResult1 = prompts.AcquirePoint(vbCrLf & "Pick label location:")

        If PromptPointResult1.Value.IsEqualTo(Point3) Then
            Return SamplerStatus.NoChange
        Else
            Point3 = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean

        Dim Arc3d As New CircularArc3d(Point1, Point2, Point3)
        Dim Circle1 As New Circle(Arc3d.Center, Vector3d.ZAxis, Point1.GetVectorTo(Arc3d.Center).Length)


        Dim Line1 As New Line(Circle1.Center, Point3)
        Dim Scale_f As Double = (((Circle1.Radius + rad_small) ^ 2 - rad_small ^ 2) ^ 0.5) / Circle1.Radius
        Line1.TransformBy(Matrix3d.Scaling(Scale_f, Circle1.Center))



        Dim PointT As New Point3d
        PointT = Line1.EndPoint
        Dim PointA As New Point3d
        PointA = Line1.GetPointAtDist(Line1.Length - rad_small)
        Dim LinieR As New Line(PointT, PointA)

        Dim LinieL As New Line
        LinieL = LinieR.Clone
        LinieL.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, PointT))
        LinieR.TransformBy(Matrix3d.Rotation(-PI / 2, Vector3d.ZAxis, PointT))


        Dim Circle2 As New Circle(LinieR.EndPoint, Vector3d.ZAxis, rad_small)

        Dim Circle3 As New Circle(LinieL.EndPoint, Vector3d.ZAxis, rad_small)

        Dim PtI1 As New Point3d
        Dim PtI2 As New Point3d

        Dim Colint1 As New Point3dCollection
        Circle1.IntersectWith(Circle2, Intersect.OnBothOperands, Colint1, IntPtr.Zero, IntPtr.Zero)
        If Colint1.Count > 0 Then
            PtI1 = Colint1(0)
        End If

        Dim Colint2 As New Point3dCollection
        Circle1.IntersectWith(Circle3, Intersect.OnBothOperands, Colint2, IntPtr.Zero, IntPtr.Zero)
        If Colint2.Count > 0 Then
            PtI2 = Colint2(0)
        End If
        If Colint1.Count > 0 And Colint2.Count > 0 Then
            If IsNothing(PtI1) = False And IsNothing(PtI2) = False Then



                Dim AngleStart As Double = GET_Bearing_rad(Circle2.Center.X, Circle2.Center.Y, PtI1.X, PtI1.Y)
                Dim AngleEnd As Double = GET_Bearing_rad(Circle2.Center.X, Circle2.Center.Y, PointT.X, PointT.Y)
                Dim Arc1 As New Arc(Circle2.Center, Rad_small, AngleStart, AngleEnd)

                AngleStart = GET_Bearing_rad(Circle3.Center.X, Circle3.Center.Y, PtI2.X, PtI2.Y)
                AngleEnd = GET_Bearing_rad(Circle3.Center.X, Circle3.Center.Y, PointT.X, PointT.Y)
                Dim Arc2 As New Arc(Circle3.Center, Rad_small, AngleEnd, AngleStart)

                Dim PointB3 As New Point3d
                Dim PointB4 As New Point3d
                PointB3 = Point1
                PointB4 = Point2

                If PointB3.GetVectorTo(Circle2.Center).Length < PointB3.GetVectorTo(Circle3.Center).Length Then
                    Dim T As New Point3d
                    T = PointB3
                    PointB3 = PointB4
                    PointB4 = T
                End If

                AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI2.X, PtI2.Y)
                AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB3.X, PointB3.Y)
                Dim Arc3 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)


                AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI1.X, PtI1.Y)
                AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB4.X, PointB4.Y)
                Dim Arc4 As New Arc(Circle1.Center, Circle1.Radius, AngleStart, AngleEnd)

                Dim Poly123 As New Polyline
                Dim b0, b1, b2, b3 As Double

                b0 = Tan(Arc3.TotalAngle / 4)
                b1 = -Tan(Arc2.TotalAngle / 4)
                b2 = -Tan(Arc1.TotalAngle / 4)
                b3 = Tan(Arc4.TotalAngle / 4)

                If Arc3.Length > Arr_len Then
                    Dim PtArr3 As New Point3d
                    PtArr3 = Arc3.GetPointAtDist(Arr_len)


                    AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr3.X, PtArr3.Y)
                    AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB3.X, PointB3.Y)
                    Dim Arc31 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                    Dim b01 As Double = Tan(Arc31.TotalAngle / 4)


                    Poly123.AddVertexAt(0, New Point2d(PointB3.X, PointB3.Y), b01, 0, Arr_width)

                    AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI2.X, PtI2.Y)
                    AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr3.X, PtArr3.Y)
                    Dim Arc32 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                    Dim b02 As Double = Tan(Arc32.TotalAngle / 4)


                    Poly123.AddVertexAt(1, New Point2d(PtArr3.X, PtArr3.Y), b02, 0, 0)
                    Poly123.AddVertexAt(2, New Point2d(PtI2.X, PtI2.Y), b1, 0, 0)
                    Poly123.AddVertexAt(3, New Point2d(PointT.X, PointT.Y), b2, 0, 0)

                    If Arc4.Length > Arr_len Then
                        Dim PtArr4 As New Point3d
                        PtArr4 = Arc4.GetPointAtDist(Arc4.Length - Arr_len)

                        AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr4.X, PtArr4.Y)
                        AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtI1.X, PtI1.Y)
                        Dim Arc41 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                        Dim b41 As Double = Tan(Arc41.TotalAngle / 4)
                        Poly123.AddVertexAt(4, New Point2d(PtI1.X, PtI1.Y), b41, 0, 0)

                        AngleStart = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PointB4.X, PointB4.Y)
                        AngleEnd = GET_Bearing_rad(Circle1.Center.X, Circle1.Center.Y, PtArr4.X, PtArr4.Y)
                        Dim Arc42 As New Arc(Circle1.Center, Circle1.Radius, AngleEnd, AngleStart)
                        Dim b42 As Double = Tan(Arc42.TotalAngle / 4)

                        Poly123.AddVertexAt(5, New Point2d(PtArr4.X, PtArr4.Y), b42, Arr_width, 0)

                        Poly123.AddVertexAt(6, New Point2d(PointB4.X, PointB4.Y), 0, 0, 0)

                        draw.Geometry.Polyline(Poly123, 0, 6)
                    End If
                End If
            End If
        End If
    End Function
End Class

Public Class jig_Mtext_class
    Inherits EntityJig
    Dim Mtext1 As MText
    Dim Location_point As New Point3d
    Dim Jig_options As JigPromptPointOptions
    Sub New(ByVal Mtext_entity As MText, ByVal TextH As Double, ByVal Textrot As Double, ByVal Content As String)
        MyBase.New(Mtext_entity)
        Mtext1 = Mtext_entity
        Mtext1.Contents = Content
        Mtext1.Rotation = Textrot
        Mtext1.TextHeight = TextH
        Mtext1.Attachment = AttachmentPoint.MiddleCenter
    End Sub
    Function BeginJig() As PromptPointResult
        If Jig_options Is Nothing Then
            Jig_options = New JigPromptPointOptions
            Jig_options.Message = vbLf & "Specify insertion point:"
        End If
        Dim editor1 As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim PromptResult1 As PromptResult = editor1.Drag(Me)
        Do
            Select Case PromptResult1.Status
                Case PromptStatus.OK
                    Return PromptResult1
                    Exit Do
                Case PromptStatus.None
                    Return PromptResult1
                    Exit Do
                Case PromptStatus.Other
                    Return PromptResult1
                    Exit Do
            End Select
        Loop While PromptResult1.Status <> PromptStatus.Cancel
        Return Nothing
    End Function

    Protected Overrides Function Sampler(ByVal prompts As JigPrompts) As SamplerStatus
        Dim Result1 As PromptPointResult = prompts.AcquirePoint(Jig_options)

        If Not Result1.Value.IsEqualTo(Location_point) = True Then
            Location_point = Result1.Value
        Else
            Return SamplerStatus.NoChange
        End If
        If Result1.Status = PromptStatus.Cancel Then
            Return SamplerStatus.Cancel
            Jig_options.Message = vbLf & "Canceled"
        Else
            Return SamplerStatus.OK
        End If
    End Function

    Protected Overrides Function Update() As Boolean

        Mtext1.Location = Location_point


        Return False
    End Function
End Class

Public Class Jig_rectangle_viewport_SHEET_CUTTER
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig

    Dim PunctM As New Point3d

    Dim Vw_width As Double
    Dim Vw_height As Double
    Dim Rotate_poly As Boolean

    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal VVw_width As Double, ByVal VVw_height As Double, ByVal Rotate As Boolean) As PromptPointResult
        Vw_width = VVw_width
        Vw_height = VVw_height
        Rotate_poly = Rotate
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptPointResult1 = ed.Drag(Me)

        Do
            Select Case PromptPointResult1.Status
                Case PromptStatus.OK
                    Return PromptPointResult1
                    Exit Do
            End Select
        Loop While PromptPointResult1.Status <> PromptStatus.Cancel
        Return PromptPointResult1
    End Function

    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus

        PromptPointResult1 = prompts.AcquirePoint(vbLf & "Specify target point for viewport:")


        If PromptPointResult1.Value.IsEqualTo(PunctM) Then
            Return SamplerStatus.NoChange
        Else
            PunctM = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
        Dim Vw_scale As Double = Sheet_cutter_form.Vw_scale
        Dim Rotation1 As Double = Sheet_cutter_form.Rotatie

        Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Poly1.AddVertexAt(0, New Point2d(PunctM.X - 0.5 * Vw_width / Vw_scale, PunctM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(1, New Point2d(PunctM.X + 0.5 * Vw_width / Vw_scale, PunctM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(2, New Point2d(PunctM.X + 0.5 * Vw_width / Vw_scale, PunctM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(3, New Point2d(PunctM.X - 0.5 * Vw_width / Vw_scale, PunctM.Y + 0.5 * Vw_height / Vw_scale), 0, 0, 0)
        Poly1.AddVertexAt(4, New Point2d(PunctM.X - 0.5 * Vw_width / Vw_scale, PunctM.Y - 0.5 * Vw_height / Vw_scale), 0, 0, 0)

        If Rotate_poly = True Then
            Poly1.TransformBy(Matrix3d.Rotation(Rotation1, Vector3d.ZAxis, PunctM))
        End If


        draw.Geometry.Polyline(Poly1, 0, 4)



    End Function
End Class


Public Class Jig_rectangle_viewport_ALIGNMENT_CUTTER
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig

    Dim Punct2 As New Point3d

    Dim Vw_width As Double
    Dim Vw_height As Double
    Dim Punct1 As New Point3d

    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal VVw_width As Double, ByVal VVw_height As Double, ByVal Start_pt As Point3d) As PromptPointResult
        Vw_width = VVw_width
        Vw_height = VVw_height
        Punct1 = Start_pt

        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptPointResult1 = ed.Drag(Me)

        Do
            Select Case PromptPointResult1.Status
                Case PromptStatus.OK
                    Return PromptPointResult1
                    Exit Do
            End Select
        Loop While PromptPointResult1.Status <> PromptStatus.Cancel
        Return PromptPointResult1
    End Function

    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus

        PromptPointResult1 = prompts.AcquirePoint(vbLf & "Specify 2nd point for viewport:")


        If PromptPointResult1.Value.IsEqualTo(Punct2) Then
            Return SamplerStatus.NoChange
        Else
            Punct2 = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
        Dim Vw_scale As Double = 1
        'Dim Rotation1 As Double = GET_Bearing_rad(Punct1.X, Punct1.Y, Punct2.X, Punct2.Y)

        Dim Line1 As New Line(Punct1, Punct2)


        If Line1.Length >= Vw_height / Vw_scale Then

            Dim Line2 As New Line(Punct1, Line1.GetPointAtDist(Vw_height / Vw_scale))
            Line2.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Punct1))

            Dim Pointm As New Point3d((Punct1.X + Line2.EndPoint.X) / 2, (Punct1.Y + Line2.EndPoint.Y) / 2, 0)

            Line2.TransformBy(Matrix3d.Displacement(Pointm.GetVectorTo(Punct1)))
            Dim Pt1 As New Point3d
            Pt1 = Line2.StartPoint
            Dim Pt2 As New Point3d
            Pt2 = Line2.EndPoint
            Line2.TransformBy(Matrix3d.Displacement(Punct1.GetVectorTo(Punct2)))

            Dim Pt4 As New Point3d
            Pt4 = Line2.StartPoint
            Dim Pt3 As New Point3d
            Pt3 = Line2.EndPoint

            Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
            Poly1.AddVertexAt(0, New Point2d(Pt1.X, Pt1.Y), 0, 0, 0)
            Poly1.AddVertexAt(1, New Point2d(Pt2.X, Pt2.Y), 0, 0, 0)
            Poly1.AddVertexAt(2, New Point2d(Pt3.X, Pt3.Y), 0, 0, 0)
            Poly1.AddVertexAt(3, New Point2d(Pt4.X, Pt4.Y), 0, 0, 0)
            Poly1.AddVertexAt(4, New Point2d(Pt1.X, Pt1.Y), 0, 0, 0)


            'Poly1.TransformBy(Matrix3d.Rotation(Rotation1, Vector3d.ZAxis, Punct1))



            draw.Geometry.Polyline(Poly1, 0, 4)

        Else







        End If




    End Function
End Class


Public Class Draw_mleader
  Inherits EntityJig
    Dim Mleader1 As MLeader
    Dim Punct1 As New Point3d
    Dim Punct2 As New Point3d
    Dim Nr2 As Integer

    Dim Jig_options As JigPromptPointOptions
    Sub New(ByVal ML1 As MLeader, ByVal pct0 As Point3d)
        MyBase.New(ML1)
        Punct1 = pct0

        Mleader1 = ML1
    End Sub
    Function BeginJig() As PromptPointResult
        If Jig_options Is Nothing Then
            Jig_options = New JigPromptPointOptions
            Jig_options.Message = vbLf & "Specify third point:"
        End If
        Dim editor1 As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim PromptResult1 As PromptResult = editor1.Drag(Me)
        Do
            Select Case PromptResult1.Status
                Case PromptStatus.OK
                    Return PromptResult1
                    Exit Do
                Case PromptStatus.None
                    Return PromptResult1
                    Exit Do
                Case PromptStatus.Other
                    Return PromptResult1
                    Exit Do
            End Select
        Loop While PromptResult1.Status <> PromptStatus.Cancel
        Return Nothing
    End Function

    Protected Overrides Function Sampler(ByVal prompts As JigPrompts) As SamplerStatus
        Dim Result1 As PromptPointResult = prompts.AcquirePoint(Jig_options)
        If Not Result1.Value.IsEqualTo(Punct2) = True Then
            Punct2 = Result1.Value
        Else
            Return SamplerStatus.NoChange
        End If
        If Result1.Status = PromptStatus.Cancel Then
            Return SamplerStatus.Cancel
            Jig_options.Message = vbLf & "Canceled"
        Else
            Return SamplerStatus.OK
        End If
    End Function

    Protected Overrides Function Update() As Boolean
        Dim Nr1 As Integer = Mleader1.AddLeader()
        Nr2 = Mleader1.AddLeaderLine(Nr1)

        Mleader1.AddFirstVertex(Nr2, New Point3d(Punct1.X, Punct1.Y, 0))
        Mleader1.AddLastVertex(Nr2, New Point3d(Punct2.X, Punct2.Y, 0))

        Dim Mtext1 As New MText

        Mtext1.Contents = "X=" & Round(Punct1.X, 2) & vbCrLf & "Y=" & Round(Punct1.Y, 2)
        Mleader1.LeaderLineType = LeaderType.StraightLeader
        Mleader1.ContentType = ContentType.MTextContent
        Mtext1.ColorIndex = 0
        Mtext1.TextHeight = 2.5
        Mleader1.MText = Mtext1
        Mleader1.LandingGap = 2
        Mleader1.ArrowSize = 2.5
        Mleader1.DoglegLength = 2
        Return False
    End Function
End Class



Public Class Jig_rectangle_viewport_SHEET_CUTTER_manual_pt2
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig

    Dim Poly_cl As Polyline
    Dim Punct1 As New Point3d
    Dim Punct2 As New Point3d

    Dim Vw_width As Double
    Dim Vw_height As Double
    Dim Vw_scale As Double

    Dim Max_dist1 As Double
    Dim Min_dist1 As Double

    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal VVw_Scale As Double, ByVal VVw_width As Double, ByVal VVw_height As Double, ByVal Poly2D As Polyline, ByVal Pt1 As Point3d, ByVal min_dist As Double, ByVal max_dist As Double) As PromptPointResult
        Vw_width = VVw_width
        Vw_height = VVw_height
        Vw_scale = VVw_Scale
        Punct1 = Pt1
        Poly_cl = Poly2D
        Max_dist1 = max_dist
        Min_dist1 = min_dist

        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptPointResult1 = ed.Drag(Me)

        Do
            Select Case PromptPointResult1.Status
                Case PromptStatus.OK
                    Return PromptPointResult1
                    Exit Do
            End Select
        Loop While PromptPointResult1.Status <> PromptStatus.Cancel
        Return PromptPointResult1
    End Function

    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus

        PromptPointResult1 = prompts.AcquirePoint(vbLf & "Please pick end location:")


        If PromptPointResult1.Value.IsEqualTo(Punct2) Then
            Return SamplerStatus.NoChange
        Else
            Punct2 = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean



        Dim Pt_on_poly1 As Point3d
        Pt_on_poly1 = Poly_cl.GetClosestPointTo(Punct1, Vector3d.ZAxis, False)
        Dim Pt_on_poly2 As Point3d
        Pt_on_poly2 = Poly_cl.GetClosestPointTo(Punct2, Vector3d.ZAxis, False)

        Dim dist1 As Double = Pt_on_poly1.DistanceTo(Pt_on_poly2)

        If dist1 > Min_dist1 And dist1 <= Max_dist1 Then
            Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
            Poly1 = creaza_rectangle_viewport(Pt_on_poly1, Pt_on_poly2)

            draw.Geometry.Polyline(Poly1, 0, 4)
        End If





    End Function


    Private Function creaza_rectangle_viewport(ByVal Point1 As Point3d, ByVal Point2 As Point3d) As Polyline

        Dim Line1R As New Line(Point1, Point2)
        Dim Point_distR As New Point3d
        If Line1R.Length > Vw_height / Vw_scale Then
            Point_distR = Line1R.GetPointAtDist(Vw_height / Vw_scale)
            Line1R.EndPoint = Point_distR
        Else
            Line1R.TransformBy(Matrix3d.Scaling((Vw_height / Vw_scale) / Line1R.Length, Line1R.StartPoint))
        End If

        Line1R.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Point1))
        Dim Point_middler As New Point3d((Point1.X + Line1R.EndPoint.X) / 2, (Point1.Y + Line1R.EndPoint.Y) / 2, 0)

        Line1R.TransformBy(Matrix3d.Displacement(Point_middler.GetVectorTo(Point1)))
        Dim Pt1r As New Point3d
        Pt1r = Line1R.StartPoint
        Dim Pt2r As New Point3d
        Pt2r = Line1R.EndPoint
        Line1R.TransformBy(Matrix3d.Displacement(Point1.GetVectorTo(Point2)))

        Dim Pt4r As New Point3d
        Pt4r = Line1R.StartPoint
        Dim Pt3r As New Point3d
        Pt3r = Line1R.EndPoint

        Dim Poly1r As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Poly1r.AddVertexAt(0, New Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0)
        Poly1r.AddVertexAt(1, New Point2d(Pt2r.X, Pt2r.Y), 0, 0, 0)
        Poly1r.AddVertexAt(2, New Point2d(Pt3r.X, Pt3r.Y), 0, 0, 0)
        Poly1r.AddVertexAt(3, New Point2d(Pt4r.X, Pt4r.Y), 0, 0, 0)
        Poly1r.AddVertexAt(4, New Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0)



        Return Poly1r

    End Function

End Class

Public Class Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig


    Private Base_point As Point3d
    Private New_position As Point3d
    Private mEntities As List(Of Entity)

    Private Linie1 As Line


    Public Sub New(ByVal basePt As Point3d, Line1 As Line)
        Base_point = basePt.TransformBy(UCS)
        mEntities = New List(Of Entity)()
        Linie1 = Line1
    End Sub


    Public Property Base() As Point3d
        Get
            Return New_position
        End Get
        Set(value As Point3d)
            New_position = value
        End Set
    End Property

    Public Property Location() As Point3d
        Get
            Return New_position
        End Get
        Set(value As Point3d)
            New_position = value
        End Set
    End Property

    Public ReadOnly Property UCS() As Matrix3d
        Get
            Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem
        End Get
    End Property

    Protected Overrides Function Sampler(prompts As JigPrompts) As SamplerStatus
        Dim prOptions1 As New JigPromptPointOptions(vbLf & "New location:")
        prOptions1.UseBasePoint = False

        Dim prResult1 As PromptPointResult = prompts.AcquirePoint(prOptions1)
        If prResult1.Status = PromptStatus.Cancel OrElse prResult1.Status = PromptStatus.[Error] Then
            Return SamplerStatus.Cancel
        End If

        If Not New_position.IsEqualTo(prResult1.Value, New Tolerance(0.000000001, 0.000000001)) Then
            New_position = prResult1.Value
            Return SamplerStatus.OK
        Else
            Return SamplerStatus.NoChange
        End If
    End Function

    Public Sub AddEntity(ent As Entity)
        mEntities.Add(ent)
    End Sub

    Public Sub TransformEntities()
        Dim Move_matrix As Matrix3d = Matrix3d.Displacement(Base_point.GetVectorTo(Linie1.GetClosestPointTo(New_position, Vector3d.ZAxis, True)))

        For Each ent As Entity In mEntities
            ent.TransformBy(Move_matrix)
        Next
    End Sub
    Protected Overrides Function WorldDraw(draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
        Dim Move_matrix As Matrix3d = Matrix3d.Displacement(Base_point.GetVectorTo(Linie1.GetClosestPointTo(New_position, Vector3d.ZAxis, True)))

        Dim geo As Autodesk.AutoCAD.GraphicsInterface.WorldGeometry = draw.Geometry
        If geo IsNot Nothing Then
            geo.PushModelTransform(Move_matrix)

            For Each ent As Entity In mEntities
                geo.Draw(ent)
            Next

            geo.PopModelTransform()
        End If

        Return True
    End Function




End Class