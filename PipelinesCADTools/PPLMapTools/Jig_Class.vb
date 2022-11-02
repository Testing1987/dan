Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.Runtime
Imports System.Math

Public Class Jig_Class
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig
    Dim Punct_off As Point3d
    Dim Punct_jig1 As Point3d
    Dim Punct_jig2 As Point3d
    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal Punct1 As Point3d, ByVal Punct2 As Point3d) As PromptPointResult
        Punct_jig1 = Punct1
        Punct_jig2 = Punct2

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

        PromptPointResult1 = prompts.AcquirePoint(vbLf & "Specify dimension position:")


        If PromptPointResult1.Value.IsEqualTo(Punct_off) Then
            Return SamplerStatus.NoChange
        Else
            Punct_off = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean


        Dim Linie1 As New Line
        Linie1.StartPoint = Punct_jig1
        Linie1.EndPoint = Punct_jig2

        Dim Punct_per As Point3d
        Punct_per = Linie1.GetClosestPointTo(Punct_off, Vector3d.ZAxis, True)
        Dim Linie2 As New Line
        Linie2.StartPoint = Punct_per
        Linie2.EndPoint = Punct_off
        Linie2.TransformBy(Matrix3d.Displacement(Punct_per.GetVectorTo(Punct_jig1)))
        Dim pct2 As Point3d = Linie2.EndPoint
        Linie2.TransformBy(Matrix3d.Displacement(Punct_jig1.GetVectorTo(Punct_jig2)))
        Dim pct3 As Point3d = Linie2.EndPoint

        Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Poly1.AddVertexAt(0, New Point2d(Punct_jig1.X, Punct_jig1.Y), 0, 0, 0)
        Poly1.AddVertexAt(1, New Point2d(pct2.X, pct2.Y), 0, 0, 0)
        Poly1.AddVertexAt(2, New Point2d(pct3.X, pct3.Y), 0, 0, 0)
        Poly1.AddVertexAt(3, New Point2d(Punct_jig2.X, Punct_jig2.Y), 0, 0, 0)
        draw.Geometry.Polyline(Poly1, 0, 3)


        'draw.Geometry.WorldLine(Arrow_start1, Arrow_start)


    End Function

   

End Class