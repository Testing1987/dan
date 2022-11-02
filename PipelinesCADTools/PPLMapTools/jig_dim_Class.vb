Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices

Public Class jig_dim_Class
    Inherits EntityJig
    Dim Aligned_dim As AlignedDimension
    Dim Punct3 As New Point3d
    Dim Jig_options As JigPromptPointOptions
    Sub New(ByVal AD As AlignedDimension, ByVal punct1 As Point3d, ByVal punct2 As Point3d)
        MyBase.New(AD)
        AD.XLine1Point = punct1
        AD.XLine2Point = punct2
        Aligned_dim = AD
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
        If Not Result1.Value.IsEqualTo(Punct3) = True Then
            Punct3 = Result1.Value
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
        Aligned_dim.DimLinePoint = Punct3
        Return False
    End Function
End Class





