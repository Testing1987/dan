
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices

Public Class Jig_highlight_poly_Class
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig

    Dim PunctM As Point3d


    Dim PolyHighlight As Polyline

    Dim PromptPointResult1 As PromptPointResult
    Function StartJig(ByVal Pollly1 As Polyline) As PromptPointResult
        PolyHighlight = Pollly1

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

        PromptPointResult1 = prompts.AcquirePoint(vbLf & "Specify insertion point:")


        If PromptPointResult1.Value.IsEqualTo(PunctM) Then
            Return SamplerStatus.NoChange
        Else
            PunctM = PromptPointResult1.Value
            Return SamplerStatus.OK
        End If


    End Function

    Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean


        Dim Poly2 As New Autodesk.AutoCAD.DatabaseServices.Polyline



        For i = 0 To PolyHighlight.NumberOfVertices - 1
            Poly2.AddVertexAt(i, PolyHighlight.GetPoint2dAt(i), 0, 10, 10)
        Next


        draw.Geometry.Polyline(Poly2, 0, PolyHighlight.NumberOfVertices - 1)

    End Function
End Class





