Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class Commands_class_PLATT
    <CommandMethod("CPU1")> _
    Public Sub gaseste_nr_serial_al_disc_C()
        Dim disk As New Management.ManagementObject("Win32_LogicalDisk.DeviceID=""C:""")
        Dim diskPropertyB As Management.PropertyData = disk.Properties("VolumeSerialNumber")
        MsgBox(diskPropertyB.Value.ToString())
    End Sub
    <CommandMethod("PLATGEN")> _
    Public Sub Show_platt_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Platt_Generator_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Platt_Generator_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("PPL_RESIDENTIAL")> _
    Public Sub Show_SHEET_CUTTER_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Sheet_cutter_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Sheet_cutter_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("sheet_cutter")> _
    Public Sub Show_alignment_cutter_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Alignment_Sheet_cutter Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Alignment_Sheet_cutter
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("PPL_DST")> _
    Public Sub Show_File_duplicator_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is File_duplicator_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New File_duplicator_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("LL1")> _
    Public Sub ARCLEADER_2()
        ' arc leader fara leader form
        If isSECURE() = False Then Exit Sub
        Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

        Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near

        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
            Dim Inaltimea1 As Double = 1000 * 0.0625

            Dim Tilemode1 As Integer = Application.GetSystemVariable("TILEMODE")
            Dim CVport1 As Integer = Application.GetSystemVariable("CVPORT")

            If (Tilemode1 = 0 And Not CVport1 = 1) Or Tilemode1 = 1 Then
                For i = System.Windows.Forms.Application.OpenForms.Count - 1 To 0 Step -1
                    Dim forma3 As System.Windows.Forms.Form = System.Windows.Forms.Application.OpenForms(i)
                    If TypeOf forma3 Is Platt_Generator_form Then
                        For Each CTRL1 As Windows.Forms.Control In forma3.Controls
                            If TypeOf CTRL1 Is Windows.Forms.TabControl Then
                                Dim Tab1 As Windows.Forms.TabControl = CTRL1
                                For Each Tab2 As Windows.Forms.TabPage In Tab1.TabPages
                                    If Tab2.Name = "TabPage_dwg_setup" Then
                                        For Each CTRL0 As Windows.Forms.Control In Tab2.Controls
                                            If TypeOf CTRL0 Is Windows.Forms.Panel Then
                                                Dim Pan1 As Windows.Forms.Panel = CTRL0
                                                If Pan1.Name = "Panel_SCALE_SELECTION" Then
                                                    For Each ctrl2 As Windows.Forms.Control In Pan1.Controls
                                                        Dim Radiob2 As Windows.Forms.RadioButton
                                                        If TypeOf ctrl2 Is Windows.Forms.RadioButton Then
                                                            Radiob2 = ctrl2
                                                            If Radiob2.Checked = True Then
                                                                If IsNumeric(Replace(Radiob2.Name, "RadioButton", "")) = True Then
                                                                    Inaltimea1 = Inaltimea1 * CInt(Replace(Radiob2.Name, "RadioButton", ""))
                                                                End If
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If

            Dim Jig1 As New Draw_JIG1
            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Pick first point : ")
            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            PP1.AllowNone = False
            Point1 = Editor1.GetPoint(PP1)

            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Point2 = Jig1.StartJig(Point1.Value, Inaltimea1)

            If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim Jig2 As New Draw_JIG2
            Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Point3 = Jig2.StartJig(Point1.Value, Point2.Value, Inaltimea1)
            If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim x0 As Double = Point1.Value.X
            Dim y0 As Double = Point1.Value.Y
            Dim x2 As Double = Point2.Value.X
            Dim y2 As Double = Point2.Value.Y
            Dim x3 As Double = Point3.Value.X
            Dim y3 As Double = Point3.Value.Y
            Dim Wdth1 As Double = (Inaltimea1) / 1250

            Dim x1, y1 As Double
            x1 = x1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)
            y1 = y1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)

            Dim Bulge1 As Double
            Bulge1 = Bulge_for_arc_leader(x0, y0, x2, y2, x3, y3, Inaltimea1)

            Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                BTrecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                Poly1.AddVertexAt(0, New Point2d(x0, y0), 0, 0, Wdth1)
                Poly1.AddVertexAt(1, New Point2d(x1, y1), Bulge1, 0, 0)
                Poly1.AddVertexAt(2, New Point2d(x3, y3), 0, 0, 0)
                BTrecord.AppendEntity(Poly1)
                Trans2.AddNewlyCreatedDBObject(Poly1, True)
                Trans2.Commit()
            End Using

        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
    End Sub


    <CommandMethod("LL200")> _
    Public Sub ARCLEADER_200()
        ' arc leader fara leader form
        If isSECURE() = False Then Exit Sub
        Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

        Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near

        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Current_UCS_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem

            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
            Dim Inaltimea1 As Double = 1000 * 0.0625 * 200



            Dim Jig1 As New Draw_JIG1
            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Pick first point : ")
            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            PP1.AllowNone = False
            Point1 = Editor1.GetPoint(PP1)

            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Point2 = Jig1.StartJig(Point1.Value.TransformBy(Current_UCS_matrix), Inaltimea1)

            If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim Jig2 As New Draw_JIG2
            Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Point3 = Jig2.StartJig(Point1.Value.TransformBy(Current_UCS_matrix), Point2.Value, Inaltimea1)
            If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim x0 As Double = Point1.Value.TransformBy(Current_UCS_matrix).X
            Dim y0 As Double = Point1.Value.TransformBy(Current_UCS_matrix).Y
            Dim x2 As Double = Point2.Value.X
            Dim y2 As Double = Point2.Value.Y
            Dim x3 As Double = Point3.Value.X
            Dim y3 As Double = Point3.Value.Y
            Dim Wdth1 As Double = (Inaltimea1) / 1250


            Dim x1, y1 As Double
            x1 = x1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)
            y1 = y1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)

            Dim Bulge1 As Double
            Bulge1 = Bulge_for_arc_leader(x0, y0, x2, y2, x3, y3, Inaltimea1)

            Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                BTrecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                Poly1.AddVertexAt(0, New Point2d(x0, y0), 0, 0, Wdth1)
                Poly1.AddVertexAt(1, New Point2d(x1, y1), Bulge1, 0, 0)
                Poly1.AddVertexAt(2, New Point2d(x3, y3), 0, 0, 0)
                BTrecord.AppendEntity(Poly1)
                Trans2.AddNewlyCreatedDBObject(Poly1, True)
                Trans2.Commit()
            End Using

        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
    End Sub


    <CommandMethod("LL100")> _
    Public Sub ARCLEADER_100()
        ' arc leader fara leader form
        If isSECURE() = False Then Exit Sub
        Dim OLD_OSnap As Integer = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

        Dim NEW_OSnap As Integer = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near

        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

        Try
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Database1 = ThisDrawing.Database
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim Current_UCS_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem

            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
            Dim Inaltimea1 As Double = 1000 * 0.0625 * 100



            Dim Jig1 As New Draw_JIG1
            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Pick first point : ")
            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            PP1.AllowNone = False
            Point1 = Editor1.GetPoint(PP1)

            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Point2 = Jig1.StartJig(Point1.Value.TransformBy(Current_UCS_matrix), Inaltimea1)

            If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim Jig2 As New Draw_JIG2
            Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Point3 = Jig2.StartJig(Point1.Value.TransformBy(Current_UCS_matrix), Point2.Value, Inaltimea1)
            If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
                Exit Sub
            End If

            Dim x0 As Double = Point1.Value.TransformBy(Current_UCS_matrix).X
            Dim y0 As Double = Point1.Value.TransformBy(Current_UCS_matrix).Y
            Dim x2 As Double = Point2.Value.X
            Dim y2 As Double = Point2.Value.Y
            Dim x3 As Double = Point3.Value.X
            Dim y3 As Double = Point3.Value.Y
            Dim Wdth1 As Double = (Inaltimea1) / 1250


            Dim x1, y1 As Double
            x1 = x1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)
            y1 = y1_for_arc_leader(x0, y0, x2, y2, Inaltimea1)

            Dim Bulge1 As Double
            Bulge1 = Bulge_for_arc_leader(x0, y0, x2, y2, x3, y3, Inaltimea1)

            Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                BTrecord = Trans2.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)
                Dim Poly1 As New Autodesk.AutoCAD.DatabaseServices.Polyline
                Poly1.AddVertexAt(0, New Point2d(x0, y0), 0, 0, Wdth1)
                Poly1.AddVertexAt(1, New Point2d(x1, y1), Bulge1, 0, 0)
                Poly1.AddVertexAt(2, New Point2d(x3, y3), 0, 0, 0)
                BTrecord.AppendEntity(Poly1)
                Trans2.AddNewlyCreatedDBObject(Poly1, True)
                Trans2.Commit()
            End Using

        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap)
    End Sub


    <CommandMethod("ListFrozenLayers", CommandFlags.NoTileMode)> _
    Sub ListFrozenLayersMethod()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        Dim ed As Editor = doc.Editor

        Dim db As Database = doc.Database



        Dim peo As New PromptEntityOptions("Select a viewport : ")

        peo.SetRejectMessage("Sorry, Not a viewport")

        peo.AddAllowedClass(GetType(Viewport), True)

        Dim per As PromptEntityResult = ed.GetEntity(peo)

        If per.Status <> PromptStatus.OK Then

            Return

        End If

        Dim vpid As ObjectId = per.ObjectId



        Using tr As Transaction = db.TransactionManager.StartTransaction()

            Dim vp As Viewport = tr.GetObject(vpid, OpenMode.ForWrite)

            Dim resBuf As ResultBuffer
            resBuf = vp.XData()
            If resBuf Is Nothing Then
                ed.WriteMessage("No layers frozen in this viewport")
            Else
                For Each tv As TypedValue In resBuf
                    Dim typeCode As Short
                    typeCode = tv.TypeCode

                    If typeCode = 1003 Then
                        ed.WriteMessage(String.Format("{0}{1}", Environment.NewLine, tv.Value.ToString))
                    End If
                Next
            End If



            tr.Commit()

        End Using

    End Sub


End Class
