Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Gis.Map
Imports Autodesk.Gis.Map.ObjectData
Imports Autodesk.Gis.Map.Constants
Imports Autodesk.Gis.Map.Utilities
Imports Autodesk.AutoCAD.EditorInput

Module Functions
    Public Function isSECURE() As Boolean
        Dim IsOK As Boolean = False
        Try

            Dim disk As New Management.ManagementObject("Win32_LogicalDisk.DeviceID=""C:""")
            Dim diskPropertyB As Management.PropertyData = disk.Properties("VolumeSerialNumber")
            If diskPropertyB.Value.ToString() = "8CDA6CE3" Then
                IsOK = True
                Return IsOK
                Exit Function
            End If

            If diskPropertyB.Value.ToString() = "36D79DE5" _
                Or diskPropertyB.Value.ToString() = "FEA3192C" _
                Or diskPropertyB.Value.ToString() = "B454BD5B" _
                Or diskPropertyB.Value.ToString() = "6E40460D" _
                Or diskPropertyB.Value.ToString() = "0892E01D" _
                Or diskPropertyB.Value.ToString() = "4ED21ABF" _
                Or diskPropertyB.Value.ToString() = "FEA3192C" _
                Or diskPropertyB.Value.ToString() = "56766C69" _
                Or diskPropertyB.Value.ToString() = "DA214366" _
                Or diskPropertyB.Value.ToString() = "3CF68AF2" _
                Or diskPropertyB.Value.ToString() = "36D79DE5" _
                Or diskPropertyB.Value.ToString() = "389A2249" _
                Or diskPropertyB.Value.ToString() = "44444444" _
                Or diskPropertyB.Value.ToString() = "8CD08F48" _
                Or diskPropertyB.Value.ToString() = "0E26E402" _
                Or diskPropertyB.Value.ToString() = "4A123A50" _
                Or diskPropertyB.Value.ToString() = "AED6B68E" _
                Or diskPropertyB.Value.ToString() = "98D9B617" _
                Or diskPropertyB.Value.ToString() = "56766C69" _
                Or diskPropertyB.Value.ToString() = "DA214366" _
                Or diskPropertyB.Value.ToString() = "B838FEB4" _
                Or diskPropertyB.Value.ToString() = "1AE1721C" _
                Or diskPropertyB.Value.ToString() = "CA9E6FFE" _
                Or diskPropertyB.Value.ToString() = "DE281128" _
                Or diskPropertyB.Value.ToString() = "FC7C4F1" _
                Or diskPropertyB.Value.ToString() = "B67EC134" _
                Or diskPropertyB.Value.ToString() = "4ED21ABF" _
                Or diskPropertyB.Value.ToString() = "E64DBF0A" _
                Or diskPropertyB.Value.ToString() = "561F1509" _
                Or diskPropertyB.Value.ToString() = "389A2249" _
                Or diskPropertyB.Value.ToString() = "B63AD3F6" _
                Or diskPropertyB.Value.ToString() = "xxxx" _
                Or diskPropertyB.Value.ToString() = "120E4B54" _
                Or diskPropertyB.Value.ToString() = "F6633173" _
                Or diskPropertyB.Value.ToString() = "40D6BDCB" _
                Or diskPropertyB.Value.ToString() = "8C040338" Then
                IsOK = True
                Return IsOK
                Exit Function
            End If

            Try
                If Environment.GetEnvironmentVariable("USERDNSDOMAIN").ToString.ToUpper = "HMMG.CC" Or Environment.GetEnvironmentVariable("USERDNSDOMAIN").ToString.ToLower() = "mottmac.group.int" Then
                    IsOK = True
                Else
                    If Today.Year = 2017 And Today.Month = 12 And Today.Day < 25 Then
                        'IsOK = True
                    End If
                End If

            Catch ex As System.Exception
                If Today.Year = 2017 And Today.Month <= 12 And Today.Day < 25 Then
                    'IsOK = True
                End If
            End Try

            If Today.Year = 2017 And Today.Month <= 12 And Today.Day < 25 Then


            End If

            Return IsOK
        Catch ex As Exception

            MsgBox(ex.Message)

            IsOK = False
            Return IsOK
        End Try

    End Function

    Public Sub ascunde_butoanele_pentru_forms(ByVal Form1 As Windows.Forms.Form, ByVal Colectie_Butoane_visibile As Specialized.StringCollection)

        For i = 0 To Form1.Controls.Count - 1
            If TypeOf Form1.Controls(i) Is Windows.Forms.Button Then
                If Form1.Controls(i).Visible = True Then
                    Colectie_Butoane_visibile.Add(Form1.Controls(i).Name)
                    Form1.Controls(i).Visible = False
                End If
            End If

            If TypeOf Form1.Controls(i) Is Windows.Forms.Panel Then
                Dim Panel1 As Windows.Forms.Panel = Form1.Controls(i)
                For j = 0 To Panel1.Controls.Count - 1
                    If TypeOf Panel1.Controls(j) Is Windows.Forms.Button Then
                        If Panel1.Controls(j).Visible = True Then
                            Colectie_Butoane_visibile.Add(Panel1.Controls(j).Name)
                            Panel1.Controls(j).Visible = False
                        End If
                    End If
                Next
            End If

            If TypeOf Form1.Controls(i) Is Windows.Forms.TabControl Then
                Dim Tab1 As Windows.Forms.TabControl = Form1.Controls(i)
                For j = 0 To Tab1.Controls.Count - 1
                    If TypeOf Tab1.Controls(j) Is Windows.Forms.Panel Then
                        Dim Panel1 As Windows.Forms.Panel = Tab1.Controls(j)
                        For k = 0 To Panel1.Controls.Count - 1
                            If TypeOf Panel1.Controls(k) Is Windows.Forms.Button Then
                                If Panel1.Controls(k).Visible = True Then
                                    Colectie_Butoane_visibile.Add(Panel1.Controls(k).Name)
                                    Panel1.Controls(k).Visible = False
                                End If
                            End If
                        Next
                    End If
                    If TypeOf Tab1.Controls(j) Is Windows.Forms.Button Then
                        If Tab1.Controls(j).Visible = True Then
                            Colectie_Butoane_visibile.Add(Tab1.Controls(j).Name)
                            Tab1.Controls(j).Visible = False
                        End If
                    End If


                Next


            End If

        Next
    End Sub
    Public Sub afiseaza_butoanele_pentru_forms(ByVal Form As Windows.Forms.Form, ByVal Colectie_Butoane_visibile As Specialized.StringCollection)

        If Colectie_Butoane_visibile.Count > 0 Then
            For i = 0 To Form.Controls.Count - 1
                If TypeOf Form.Controls(i) Is Windows.Forms.Button Then
                    If Colectie_Butoane_visibile.Contains(Form.Controls(i).Name) = True Then
                        Form.Controls(i).Visible = True
                    End If
                End If
                If TypeOf Form.Controls(i) Is Windows.Forms.Panel Then
                    Dim Panel1 As Windows.Forms.Panel = Form.Controls(i)
                    For j = 0 To Panel1.Controls.Count - 1
                        If TypeOf Panel1.Controls(j) Is Windows.Forms.Button Then
                            If Colectie_Butoane_visibile.Contains(Panel1.Controls(j).Name) = True Then
                                Panel1.Controls(j).Visible = True
                            End If
                        End If
                    Next
                End If

                If TypeOf Form.Controls(i) Is Windows.Forms.TabControl Then
                    Dim Tab1 As Windows.Forms.TabControl = Form.Controls(i)
                    For j = 0 To Tab1.Controls.Count - 1
                        If TypeOf Tab1.Controls(j) Is Windows.Forms.Panel Then
                            Dim Panel1 As Windows.Forms.Panel = Tab1.Controls(j)
                            For k = 0 To Panel1.Controls.Count - 1
                                If TypeOf Panel1.Controls(k) Is Windows.Forms.Button Then
                                    If Colectie_Butoane_visibile.Contains(Panel1.Controls(k).Name) = True Then
                                        Panel1.Controls(k).Visible = True
                                    End If
                                End If
                            Next
                        End If
                        If TypeOf Tab1.Controls(j) Is Windows.Forms.Button Then
                            If Colectie_Butoane_visibile.Contains(Tab1.Controls(j).Name) = True Then
                                Tab1.Controls(j).Visible = True
                            End If
                        End If
                    Next



                End If

            Next
            Colectie_Butoane_visibile.Clear()
        End If
    End Sub


    Public Function Degree_symbol() As String
        Return Chr(176)
    End Function
    Public Function TAB_symbol() As String
        Return Chr(9)
    End Function

    Public Sub Creaza_layer_with_database(ByVal Database1 As Autodesk.AutoCAD.DatabaseServices.Database, ByVal Layername1 As String, ByVal Culoare As Integer, ByVal Descriptie As String, ByVal Plot As Boolean)
        Try
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Database1.TransactionManager.StartTransaction
                Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                LayerTable1 = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If LayerTable1.Has(Layername1) = False Then
                    LayerTable1.UpgradeOpen()
                    Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord = New Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                    LayerTableRecord1.Name = Layername1
                    LayerTableRecord1.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare)
                    LayerTableRecord1.IsPlottable = Plot
                    LayerTable1.Add(LayerTableRecord1)
                    Trans1.AddNewlyCreatedDBObject(LayerTableRecord1, True)
                    LayerTableRecord1.Description = Descriptie
                End If
                Trans1.Commit()
            End Using


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Incarca_existing_layers_to_combobox(ByVal Combo_layer As Windows.Forms.ComboBox)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument



                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    Combo_layer.Items.Clear()
                    Dim Array() As String
                    Dim i As Integer = 0
                    For Each Layer_id As ObjectId In layer_table
                        Dim Layer1 As LayerTableRecord = Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If Layer1.Name.Contains("|") = False Then
                            If Layer1.Name.Contains("$") = False Then
                                ReDim Preserve Array(i)
                                Array(i) = Layer1.Name
                                i = i + 1
                            End If
                        End If

                    Next
                    System.Array.Sort(Of String)(Array)

                    For i = 0 To Array.Length - 1
                        Combo_layer.Items.Add(Array(i))
                    Next





                End Using ' asta e de la trans1

            End Using

            Combo_layer.SelectedIndex = 0
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub CREAZA_LEADER(ByVal Punct3D As Point3d, ByVal Text_content As String, ByVal Mtext_Height As Double, ByVal Rotatie As Double, ByVal Layer As String,
                             ByVal Arrow_size As Double, ByVal Dogleg As Double, ByVal LandGap As Double, ByVal ColorIndex As Integer, ByVal LeaderColor As Integer, ByVal LastVertex As Point3d)
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try

            Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                    Dim Mtext1 As New MText
                    Mtext1.Contents = Text_content
                    Mtext1.TextHeight = Mtext_Height
                    Mtext1.ColorIndex = ColorIndex
                    Mtext1.Rotation = Rotatie

                    Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem

                    Dim Mleader1 As New MLeader
                    Dim Nr1 As Integer = Mleader1.AddLeader()
                    Dim Nr2 As Integer = Mleader1.AddLeaderLine(Nr1)
                    Mleader1.AddFirstVertex(Nr2, Punct3D.TransformBy(Curent_UCS))
                    Mleader1.AddLastVertex(Nr2, LastVertex.TransformBy(Curent_UCS))

                    Mleader1.ContentType = ContentType.MTextContent
                    Mleader1.MText = Mtext1
                    Mleader1.Layer = Layer
                    Mleader1.LandingGap = LandGap
                    Mleader1.ArrowSize = Arrow_size
                    Mleader1.DoglegLength = Dogleg
                    Mleader1.LeaderLineColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, LeaderColor)


                    BTrecord.AppendEntity(Mleader1)
                    Trans1.AddNewlyCreatedDBObject(Mleader1, True)
                    Trans1.Commit()
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Function ShowObjectDataInfo(ByVal tables As Tables, ByVal id As ObjectId) As Boolean
        Dim errCode As ErrorCode = ErrorCode.OK
        Try
            Dim success As Boolean = True

            ' Get and Initialize Records
            Dim records As Records = tables.GetObjectRecords(Convert.ToUInt32(0), id, Constants.OpenMode.OpenForRead, False)

            Try
                If records.Count = 0 Then
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbNewLine + " There is no ObjectData record attached on the entity. ")
                    Return True
                End If

                Dim index As Integer = 0

                ' Iterate through all records
                Dim record As Record
                For Each record In records
                    Dim msg As String = Nothing
                    msg = String.Format(vbNewLine + "Record {0} : ", index)
                    index = index + 1
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(msg)

                    ' Get the table
                    Dim table As ObjectData.Table = tables(record.TableName)
                    Dim tableDef As FieldDefinitions = table.FieldDefinitions

                    Dim valInt As Integer = 0
                    Dim valDouble As Double = 0.0
                    Dim str As String = Nothing

                    ' Get record info
                    Dim i As Integer
                    Dim upbound As Integer = record.Count - 1
                    For i = 0 To upbound
                        Dim column As FieldDefinition = tableDef(i)
                        Dim colName As String = column.Name
                        Dim val As MapValue = record(i)

                        Select Case val.Type
                            Case Constants.DataType.Integer
                                valInt = val.Int32Value
                                msg = String.Format("{0}; ", valInt)
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(msg)

                            Case Constants.DataType.Real
                                valDouble = val.DoubleValue
                                msg = String.Format("{0}; ", valDouble)
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(msg)

                            Case Constants.DataType.Character
                                str = val.StrValue
                                msg = String.Format("{0}; ", str)
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(msg)

                            Case Constants.DataType.Point
                                Dim pt As Point3d = val.Point
                                Dim x As Double = pt.X
                                Dim y As Double = pt.Y
                                Dim z As Double = pt.Z
                                msg = String.Format("Point({0},{1},{2}); ", x, y, z)
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(msg)

                            Case Else
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbNewLine + "Wrong data type" + vbNewLine)
                                success = False
                        End Select
                    Next i
                Next
            Finally
                records.Dispose()
            End Try

            Return success
        Catch err As MapException
            errCode = CType((err.ErrorCode), Constants.ErrorCode)
            ' Deal with the exception here as your will
            Return False
        End Try
    End Function

    Public Function get_new_worksheet_from_Excel() As Microsoft.Office.Interop.Excel.Worksheet
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook
        Try
            Try
                Excel1 = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
            Catch ex1 As System.SystemException
                Excel1 = New Microsoft.Office.Interop.Excel.Application
            End Try

            If IsNothing(Excel1) = False Then
                Excel1.Visible = True
                Excel1.Workbooks.Add()
                Workbook1 = Excel1.ActiveWorkbook
                Return Workbook1.ActiveSheet
            Else
                Return Nothing
            End If
        Catch ex As System.Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function Creaza_Mleader_nou_fara_UCS_transform(ByVal Point1 As Point3d, ByVal Continut As String, ByVal text_height As Double, ByVal arrow_size As Double, ByVal Gap1 As Double, ByVal DELTA_X As Double, ByVal DELTA_Y As Double) As MLeader
        Try


            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                Dim Mtext1 As New MText
                Mtext1.TextHeight = text_height
                Mtext1.Contents = Continut
                Mtext1.ColorIndex = 0

                Dim Mleader1 As New MLeader
                Dim Nr1 As Integer = Mleader1.AddLeader()
                Dim Nr2 As Integer = Mleader1.AddLeaderLine(Nr1)
                Mleader1.AddFirstVertex(Nr2, New Point3d(Point1.X, Point1.Y, Point1.Z))
                Mleader1.AddLastVertex(Nr2, New Point3d(Point1.X + DELTA_X, Point1.Y + DELTA_Y, Point1.Z))
                Mleader1.LeaderLineType = LeaderType.StraightLeader
                Mleader1.ContentType = ContentType.MTextContent
                Mleader1.MText = Mtext1
                Mleader1.TextHeight = text_height
                Mleader1.LandingGap = Gap1
                Mleader1.ArrowSize = arrow_size
                Mleader1.DoglegLength = Gap1
                'Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader)
                'Mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader)
                Mleader1.Annotative = AnnotativeStates.False
                BTrecord.AppendEntity(Mleader1)
                Trans1.AddNewlyCreatedDBObject(Mleader1, True)
                Trans1.Commit()
                Return Mleader1
            End Using



        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Creaza_Mleader_nou_fara_UCS_transform_CU_btrecord_AND_TRANS(ByVal BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord, ByVal Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction, ByVal Point1 As Point3d, ByVal Continut As String, ByVal text_height As Double, ByVal arrow_size As Double, ByVal Gap1 As Double, ByVal DELTA_X As Double, ByVal DELTA_Y As Double) As MLeader
        Try


            Dim Mtext1 As New MText
            Mtext1.TextHeight = text_height
            Mtext1.Contents = Continut
            Mtext1.ColorIndex = 0

            Dim Mleader1 As New MLeader
            Dim Nr1 As Integer = Mleader1.AddLeader()
            Dim Nr2 As Integer = Mleader1.AddLeaderLine(Nr1)
            Mleader1.AddFirstVertex(Nr2, New Point3d(Point1.X, Point1.Y, Point1.Z))
            Mleader1.AddLastVertex(Nr2, New Point3d(Point1.X + DELTA_X, Point1.Y + DELTA_Y, Point1.Z))
            Mleader1.LeaderLineType = LeaderType.StraightLeader
            Mleader1.ContentType = ContentType.MTextContent
            Mleader1.MText = Mtext1
            Mleader1.TextHeight = text_height
            Mleader1.LandingGap = Gap1
            Mleader1.ArrowSize = arrow_size
            Mleader1.DoglegLength = Gap1

            Mleader1.Annotative = AnnotativeStates.False
            BTrecord.AppendEntity(Mleader1)
            Trans1.AddNewlyCreatedDBObject(Mleader1, True)

            Return Mleader1




        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Get_active_worksheet_from_Excel() As Microsoft.Office.Interop.Excel.Worksheet
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook
        Try
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch ex As System.SystemException
            MsgBox("No excel found")
            Excel1 = New Microsoft.Office.Interop.Excel.Application
            Excel1.Visible = True
            Excel1.Workbooks.Add()
        Finally
            'Excel1.ActiveWindow.DisplayGridlines = True
            If Excel1.Workbooks.Count = 0 Then Excel1.Workbooks.Add()
            If Excel1.Visible = False Then Excel1.Visible = True
            Workbook1 = Excel1.ActiveWorkbook
        End Try
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Return Workbook1.ActiveSheet
    End Function
    Public Function Get_active_workbook_from_Excel() As Microsoft.Office.Interop.Excel.Workbook
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook
        Try
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch ex As System.SystemException
            MsgBox("No excel found")
            Excel1 = New Microsoft.Office.Interop.Excel.Application
            Excel1.Visible = True
            Excel1.Workbooks.Add()
        Finally
            'Excel1.ActiveWindow.DisplayGridlines = True
            If Excel1.Workbooks.Count = 0 Then Excel1.Workbooks.Add()
            If Excel1.Visible = False Then Excel1.Visible = True
            Workbook1 = Excel1.ActiveWorkbook
        End Try
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Return Workbook1
    End Function
    Public Function Get_chainage_from_double(ByVal Numar As Double, ByVal Nr_dec As Integer) As String
        Dim String2, String3 As String

        Dim String_minus As String = ""

        If Numar < 0 Then
            String_minus = "-"
            Numar = -Numar
        End If

        String2 = Get_String_Rounded(Numar, Nr_dec)

        Dim Punct As Integer
        If InStr(String2, ".") = 0 Then
            Punct = 0
        Else
            Punct = 1
        End If

        If Len(String2) - Nr_dec - Punct >= 4 Then
            String3 = Left(String2, Len(String2) - 3 - Nr_dec - Punct) & "+" & Right(String2, 3 + Nr_dec + Punct)
        Else
            If Len(String2) - Nr_dec - Punct = 1 Then String3 = "0+00" & String2
            If Len(String2) - Nr_dec - Punct = 2 Then String3 = "0+0" & String2
            If Len(String2) - Nr_dec - Punct > 2 Then String3 = "0+" & String2

        End If
        Return String_minus & String3
    End Function
    Public Function Get_String_Rounded(ByVal Numar As Double, ByVal Nr_dec As Integer) As String
        Dim String1, String2, Zero, zero1 As String
        Zero = ""
        zero1 = ""

        Dim String_punct As String = ""

        If Nr_dec > 0 Then
            String_punct = "."
            For i = 1 To Nr_dec
                Zero = Zero & "0"
            Next
        End If

        Dim String_minus As String = ""

        If Numar < 0 Then
            String_minus = "-"
            Numar = -Numar
        End If

        String1 = Round(Numar, Nr_dec, MidpointRounding.AwayFromZero).ToString
        String2 = String1

        If InStr(String1, ".") = 0 Then
            String2 = String1 & String_punct & Zero
            GoTo 123
        End If

        If Len(String1) - InStr(String1, ".") - Nr_dec <> 0 Then
            For i = 1 To InStr(String1, ".") + Nr_dec - Len(String1)
                zero1 = zero1 & "0"
            Next

            String2 = String1 & zero1
        End If
123:
        Return String_minus & String2
    End Function
    Public Function Get_chainage_with_CSF(ByVal Poly3d As Curve, ByVal point_on_poly3d As Point3d, ByVal DataTable_text_zero As System.Data.DataTable, ByVal DataTable_text_325 As System.Data.DataTable) As Double
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

            Dim Parameter_picked As Double = Round(Poly3d.GetParameterAtPoint(point_on_poly3d), 3)

            Dim Parameter_start As Double = Floor(Parameter_picked)
            Dim Parameter_end As Double = Ceiling(Parameter_picked)
            If Parameter_picked = Round(Parameter_picked, 0) Then
                Parameter_start = Parameter_picked
                Parameter_end = Parameter_picked
            End If

            Dim Chainage_on_vertex As Double
            Dim Distanta_pana_la_Vertex As Double
            Dim CSF1, CSF2 As Double
            Dim Chainage_fara_csf As Double

            If DataTable_text_325.Rows.Count > 0 Then
                Dim Point_CHAINAGE As New Point3d
                Point_CHAINAGE = Poly3d.GetPointAtParameter(Parameter_start)
                Distanta_pana_la_Vertex = Point_CHAINAGE.GetVectorTo(point_on_poly3d).Length

                For i = 0 To DataTable_text_325.Rows.Count - 1
                    Dim Text1 As DBText = DataTable_text_325.Rows(i).Item("TEXT325")
                    If Point_CHAINAGE.GetVectorTo(Text1.Position.TransformBy(ThisDrawing.Editor.CurrentUserCoordinateSystem)).Length < 0.1 Then
                        Dim String1 As String = Replace(Text1.TextString, "+", "")
                        If IsNumeric(String1) = True Then
                            Chainage_on_vertex = CDbl(String1)
                            Exit For
                        End If
                    End If
                Next

            End If

            If Chainage_on_vertex = 0 Then
                Chainage_fara_csf = Poly3d.GetDistAtPoint(point_on_poly3d)
                Return Chainage_fara_csf
                Exit Function
            End If

            If Not Parameter_start = Parameter_end Then
                If DataTable_text_zero.Rows.Count > 0 Then
                    Dim Point_CHAINAGE1 As New Point3d
                    Point_CHAINAGE1 = Poly3d.GetPointAtParameter(Parameter_start)
                    Dim Point_CHAINAGE2 As New Point3d
                    Point_CHAINAGE2 = Poly3d.GetPointAtParameter(Parameter_end)

                    For i = 0 To DataTable_text_zero.Rows.Count - 1
                        Dim Text1 As DBText = DataTable_text_zero.Rows(i).Item("TEXT0")
                        Dim String1 As String = Text1.TextString
                        String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                        If IsNumeric(String1) = True Then
                            If Point_CHAINAGE1.GetVectorTo(Text1.Position.TransformBy(ThisDrawing.Editor.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                    CSF1 = CDbl(String1)
                                End If
                            End If

                            If Point_CHAINAGE2.GetVectorTo(Text1.Position.TransformBy(ThisDrawing.Editor.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                    CSF2 = CDbl(String1)
                                End If
                            End If
                        End If
                    Next
                End If
            End If


            Dim New_ch As Double
            If Not CSF1 + CSF2 = 0 And Not CSF1 = 0 And Not CSF2 = 0 Then
                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex / ((CSF1 + CSF2) / 2)
            Else
                New_ch = Chainage_on_vertex + Distanta_pana_la_Vertex
            End If
            Return New_ch
        End Using
    End Function
    Public Sub Creaza_layer(ByVal Layername1 As String, ByVal Culoare As Integer, ByVal Descriptie As String, ByVal Plot As Boolean)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database = ThisDrawing.Database
                Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                LayerTable1 = Trans2.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If LayerTable1.Has(Layername1) = False Then
                    LayerTable1.UpgradeOpen()
                    Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord = New Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                    LayerTableRecord1.Name = Layername1
                    LayerTableRecord1.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare)
                    LayerTableRecord1.IsPlottable = Plot
                    LayerTable1.Add(LayerTableRecord1)
                    Trans2.AddNewlyCreatedDBObject(LayerTableRecord1, True)
                    LayerTableRecord1.Description = Descriptie
                    Trans2.Commit()
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function GET_Bearing_rad(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

        Return New Point3d(x1, y1, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)



    End Function
    Public Function Get_chainage_feet_from_double(ByVal Numar As Double, ByVal Nr_dec As Integer) As String
        Dim String2, String3 As String

        Dim String_minus As String = ""

        If Numar < 0 Then
            String_minus = "-"
            Numar = -Numar
        End If

        String2 = Get_String_Rounded(Numar, Nr_dec)

        Dim Punct As Integer
        If InStr(String2, ".") = 0 Then
            Punct = 0
        Else
            Punct = 1
        End If

        If Len(String2) - Nr_dec - Punct >= 4 Then
            String3 = Left(String2, Len(String2) - 2 - Nr_dec - Punct) & "+" & Right(String2, 2 + Nr_dec + Punct)
        Else
            If Len(String2) - Nr_dec - Punct = 1 Then String3 = "0+0" & String2
            If Len(String2) - Nr_dec - Punct = 2 Then String3 = "0+" & String2
            If Len(String2) - Nr_dec - Punct = 3 Then String3 = Left(String2, 1) & "+" & Right(String2, 2 + Nr_dec + Punct)

        End If
        Return String_minus & String3
    End Function
    Public Function Directie_offset(ByVal ent As ObjectId, ByVal pt As Point3d) As Integer
        Using trans As Transaction = Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
            Dim cur As Curve = trans.GetObject(ent, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False)
            Dim ptOnObj As Point3d = cur.GetClosestPointTo(pt, Vector3d.ZAxis, False)
            Dim pln = New Plane(New Point3d(0, 0, 0), New Vector3d(0, 0, 1))
            Dim ptAlongObj As Point3d
            Try
                ptAlongObj = cur.GetPointAtDist(cur.GetDistAtPoint(ptOnObj) + 1)
            Catch ex As Exception
                ptAlongObj = cur.EndPoint
            End Try
            Dim vecOnObj As Vector3d = ptAlongObj - ptOnObj
            Dim vecToPoint As Vector3d = pt - ptOnObj

            Dim angAlongObj As Double = vecOnObj.AngleOnPlane(pln)

            Dim ang As Double = vecToPoint.AngleOnPlane(pln)
            If ang < angAlongObj Then ang += Math.PI * 2


            If angAlongObj + Math.PI < ang Then
                If TypeOf cur Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Return -1
                Else
                    Return 1
                End If
                Return 1
            Else
                If TypeOf cur Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Return 1
                Else
                    Return -1
                End If
            End If


        End Using

    End Function

    Public Function Get_chainage_with_CSF_from_dbtext(ByVal Curve1 As Curve, ByVal point_on_poly3d As Point3d, ByVal ColectieDBtext_csf As DBObjectCollection, ByVal ColectieDBtext_chainage As DBObjectCollection) As Double
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



            Dim Parameter_picked As Double = Curve1.GetParameterAtPoint(point_on_poly3d)

            Dim Parameter_start As Double = Floor(Parameter_picked)
            Dim Parameter_end As Double = Ceiling(Parameter_picked)
            If Parameter_picked = Round(Parameter_picked, 0) Then
                Parameter_start = Parameter_picked
                Parameter_end = Parameter_picked
            End If

            Dim Chainage_on_vertex As Double
            Dim Distanta_pana_la_Vertex As Double
            Dim CSF1 As Double = 0
            Dim CSF2 As Double = 0

            If ColectieDBtext_csf.Count > 0 Then
                Dim Point_CHAINAGE As New Point3d
                Point_CHAINAGE = Curve1.GetPointAtParameter(Parameter_start)
                Distanta_pana_la_Vertex = Point_CHAINAGE.GetVectorTo(point_on_poly3d).Length

                For i = 0 To ColectieDBtext_chainage.Count - 1
                    Dim Text1 As DBText = ColectieDBtext_chainage(i)
                    If Point_CHAINAGE.GetVectorTo(Text1.Position.TransformBy(ThisDrawing.Editor.CurrentUserCoordinateSystem)).Length < 0.1 Then
                        Dim String1 As String = Replace(Text1.TextString, "+", "")
                        If IsNumeric(String1) = True Then
                            Chainage_on_vertex = CDbl(String1)
                            Exit For
                        End If
                    End If
                Next

            Else
                Chainage_on_vertex = Curve1.GetDistanceAtParameter(Parameter_start)
                Distanta_pana_la_Vertex = Curve1.GetPointAtParameter(Parameter_start).GetVectorTo(point_on_poly3d).Length
            End If

            If Not Parameter_start = Parameter_end Then
                If ColectieDBtext_csf.Count > 0 Then
                    Dim Point_CHAINAGE1 As New Point3d
                    Point_CHAINAGE1 = Curve1.GetPointAtParameter(Parameter_start)
                    Dim Point_CHAINAGE2 As New Point3d
                    Point_CHAINAGE2 = Curve1.GetPointAtParameter(Parameter_end)

                    For i = 0 To ColectieDBtext_csf.Count - 1
                        Dim Text1 As DBText = ColectieDBtext_csf(i)
                        Dim String1 As String = Text1.TextString
                        String1 = extrage_numar_din_text_de_la_sfarsitul_textului(String1)
                        If IsNumeric(String1) = True Then
                            If Point_CHAINAGE1.GetVectorTo(Text1.Position.TransformBy(ThisDrawing.Editor.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                    CSF1 = CDbl(String1)
                                End If
                            End If

                            If Point_CHAINAGE2.GetVectorTo(Text1.Position.TransformBy(ThisDrawing.Editor.CurrentUserCoordinateSystem)).Length < 0.1 Then
                                If CDbl(String1) > 0.5 And CDbl(String1) < 1.5 Then
                                    CSF2 = CDbl(String1)
                                End If
                            End If
                        End If
                    Next
                Else
                    CSF1 = 1
                    CSF2 = 1

                End If
            End If


            Dim New_chainage_CSF As Double
            If Not CSF1 + CSF2 = 0 And Not CSF1 = 0 And Not CSF2 = 0 Then
                New_chainage_CSF = Chainage_on_vertex + Distanta_pana_la_Vertex / ((CSF1 + CSF2) / 2)
            Else
                New_chainage_CSF = Chainage_on_vertex + Distanta_pana_la_Vertex
            End If
            Return New_chainage_CSF



        End Using
    End Function

    Public Function InsertBlock_with_multiple_atributes(ByVal Nume_fisier As String, ByVal NumeBlock As String,
                                                               ByVal Insertion_point As Point3d, ByVal Scale_xyz As Double, ByVal Spatiu As BlockTableRecord,
                                                               ByVal Layer1 As String,
                                                               ByVal Colectie_nume_atribute As Specialized.StringCollection, Colectie_valori_atribute As Specialized.StringCollection) As BlockReference
        Dim dlock As DocumentLock = Nothing
        Dim BlockTable1 As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing
        Dim br As BlockReference = Nothing
        Dim id As ObjectId
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor

            'insert block and rename it
            Try
                Try
                    dlock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Catch ex As Exception
                    Dim aex As New System.Exception("Error locking document for InsertBlock: " & NumeBlock & ": ", ex)
                    Throw aex
                End Try
                BlockTable1 = trans.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
12345:
                If BlockTable1.Has(NumeBlock) = True Then
                    Block_table_record1 = trans.GetObject(BlockTable1.Item(NumeBlock), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                Else
                    Try
                        Dim Fisier1 As String = "C:\BLOCKS_Transcanada\" & Nume_fisier

                        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Fisier1) = True Then
                            Dim ThisDrawing As Document = Application.DocumentManager.MdiActiveDocument

                            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                                Using Database2 As New Database(False, False)
                                    'read block drawing do we need Lockdocument..?? Not Always..?
                                    Database2.ReadDwgFile(Fisier1, System.IO.FileShare.Read, True, Nothing)
                                    Using Trans1 As Transaction = ThisDrawing.TransactionManager.StartTransaction()

                                        Dim BlockTable2 As BlockTable = DirectCast(Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False), BlockTable)

                                        Dim idBTR As ObjectId = ThisDrawing.Database.Insert(NumeBlock, Database2, False)
                                        Trans1.Commit()
                                        GoTo 12345
                                    End Using
                                End Using
                            End Using
                        Else
                            Return br

                        End If





                    Catch e As System.Exception

                        MsgBox(e.Message)
                    End Try

                End If ' ASTA E DE LA  If bt.Has THE BLOCK

                Spatiu = trans.GetObject(Spatiu.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                'Set the Attribute Value
                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                Dim ent As Entity
                Dim Block_table_record1enum As BlockTableRecordEnumerator
                br = New BlockReference(Insertion_point, Block_table_record1.ObjectId)
                br.Layer = Layer1
                br.ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(Scale_xyz, Scale_xyz, Scale_xyz)

                Spatiu.AppendEntity(br)
                trans.AddNewlyCreatedDBObject(br, True)
                attColl = br.AttributeCollection
                Block_table_record1enum = Block_table_record1.GetEnumerator
                While Block_table_record1enum.MoveNext
                    ent = Block_table_record1enum.Current.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    If TypeOf ent Is AttributeDefinition Then
                        Dim attdef As AttributeDefinition = ent
                        Dim attref As New AttributeReference
                        attref.SetAttributeFromBlock(attdef, br.BlockTransform)

                        For i = 0 To Colectie_nume_atribute.Count - 1
                            If attref.Tag = Colectie_nume_atribute(i) Then

                                If Not Replace(Colectie_valori_atribute(i), " ", "") = "" Then
                                    If String.IsNullOrEmpty(Colectie_valori_atribute(i)) = False Then
                                        attref.TextString = Colectie_valori_atribute(i)
                                        Exit For
                                    End If
                                End If

                            End If
                        Next


                        If IsNothing(attref) = False Then
                            attColl.AppendAttribute(attref)
                            trans.AddNewlyCreatedDBObject(attref, True)
                        End If



                    End If
                End While
                trans.Commit()

            Catch ex As System.Exception
                Dim aex2 As New System.Exception("Error in inserting new block: " & Nume_fisier & ": ", ex)
                Throw aex2
            Finally
                If Not trans Is Nothing Then trans.Dispose()
                If Not dlock Is Nothing Then dlock.Dispose()
            End Try
        End Using
        Return br
    End Function


    Public Function Locatie_fisier(ByVal Locatie1_mydoc As String, ByVal Locatie2_server As String) As String

        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Locatie1_mydoc) = True Then
            Return Locatie1_mydoc
        Else
            If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Locatie2_server) = True Then

                Return Locatie2_server
            End If
        End If

    End Function
    Public Function Incarca_existing_Blocks_with_attributes_to_combobox(ByVal Combo_Blocks_with_atributes As Windows.Forms.ComboBox)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument



                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim Block_table As Autodesk.AutoCAD.DatabaseServices.BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Combo_Blocks_with_atributes.Items.Clear()

                    Combo_Blocks_with_atributes.Items.Add("")
                    For Each Block_id As ObjectId In Block_table
                        Dim Block1 As BlockTableRecord = Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If Block1.HasAttributeDefinitions = True Then
                            Combo_Blocks_with_atributes.Items.Add(Block1.Name)
                        End If
                    Next




                End Using ' asta e de la trans1

            End Using

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function

    Public Function Incarca_existing_Atributes_to_combobox(ByVal BlockName As String, ByVal Combo_atributes As Windows.Forms.ComboBox)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument



                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim Block_table As Autodesk.AutoCAD.DatabaseServices.BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    If Combo_atributes.Items.Count > 0 Then Combo_atributes.Items.Clear()
                    If Not BlockName = "" Then
                        Dim BTrecordBlock As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Block_table(BlockName), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        For Each Id1 As ObjectId In BTrecordBlock
                            Dim ent As Entity = TryCast(Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Entity)
                            If ent IsNot Nothing Then
                                Dim attDefinition1 As AttributeDefinition = TryCast(ent, AttributeDefinition)
                                If attDefinition1 IsNot Nothing Then
                                    Combo_atributes.Items.Add(attDefinition1.Tag)
                                End If
                            End If


                        Next
                    End If





                End Using ' asta e de la trans1

            End Using

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function



    Public Function Incarca_existing_Blocks_to_combobox(ByVal Combo_Blocks As Windows.Forms.ComboBox)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument



                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim Block_table As Autodesk.AutoCAD.DatabaseServices.BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Combo_Blocks.Items.Clear()

                    Combo_Blocks.Items.Add("")
                    For Each Block_id As ObjectId In Block_table
                        Dim Block1 As BlockTableRecord = Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If Block1.Name.Contains("*") = False And Block1.IsFromExternalReference = False And Block1.IsFromOverlayReference = False Then
                            Combo_Blocks.Items.Add(Block1.Name)
                        End If


                    Next




                End Using ' asta e de la trans1

            End Using

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function
    Public Function get_block_name(ByVal bref As BlockReference) As String
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim Block_table As Autodesk.AutoCAD.DatabaseServices.BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    If bref.IsDynamicBlock = True Then
                        Dim btr As BlockTableRecord = Trans1.GetObject(bref.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Return btr.Name
                    Else

                        Dim btr As BlockTableRecord = Trans1.GetObject(bref.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Return btr.Name
                    End If
                End Using ' asta e de la trans1
            End Using
        Catch ex As System.Exception
            Return ""
        End Try

    End Function



    Public Function WCS_align() As Matrix3d

        Try

            Dim Point1 As New Point3d(0, 0, 0)
            Dim Point2 As New Point3d(0, 0, 11)
            Dim Point3 As New Point3d(0, 2, 0)
            Dim ZAxis As Vector3d = Point1.GetVectorTo(Point2).GetNormal
            Dim yAxis As Vector3d = Point1.GetVectorTo(Point3).GetNormal
            'Dim yAxis As Vector3d = ZAxis.GetPerpendicularVector.GetNormal

            Dim xAxis As Vector3d = yAxis.CrossProduct(ZAxis).GetNormal

            Dim NewMatrix3d As Matrix3d = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, Point1, xAxis, yAxis, ZAxis)

            Return NewMatrix3d
        Catch
            Return Nothing
        End Try
    End Function
    Public Function UCS_align_FRONT() As Matrix3d
        Try
            Dim Point1 As New Point3d(0, 0, 0)
            Dim Point2 As New Point3d(0, -10, 0)
            Dim Point3 As New Point3d(0, 0, 10)
            Dim ZAxis As Vector3d = Point1.GetVectorTo(Point2).GetNormal
            Dim yAxis As Vector3d = Point1.GetVectorTo(Point3).GetNormal
            'Dim yAxis As Vector3d = ZAxis.GetPerpendicularVector.GetNormal
            Dim xAxis As Vector3d = yAxis.CrossProduct(ZAxis).GetNormal
            Dim NewMatrix3d As Matrix3d = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, Point1, xAxis, yAxis, ZAxis)
            Return NewMatrix3d
        Catch
            Return Nothing
        End Try
    End Function

    Public Function Sort_data_table(ByVal Datatable1 As System.Data.DataTable, ByVal Column1 As String) As System.Data.DataTable
        If Datatable1.Columns.Contains(Column1) = True Then
            Dim DataView1 As New DataView(Datatable1)
            DataView1.Sort = Column1 & " ASC"
            'MsgBox(DataView1.Count)

            Dim Data_table_temp As New System.Data.DataTable
            Data_table_temp = Datatable1.Clone
            Data_table_temp.Rows.Clear()

            Dim k As Integer = 0
            For i = 0 To DataView1.Count - 1
                Dim Data_row1 As DataRow
                Data_row1 = DataView1(i).Row
                Data_table_temp.Rows.Add()
                For j = 0 To Datatable1.Columns.Count - 1
                    Data_table_temp.Rows(k).Item(j) = Data_row1.Item(j)
                Next
                k = k + 1
            Next
            Return Data_table_temp
        End If




    End Function

    Public Function Sort_data_table_2_columns(ByVal Datatable1 As System.Data.DataTable, ByVal Column1 As String, ByVal SPACE_ASC_DESC1 As String, ByVal Column2 As String, ByVal SPACE_ASC_DESC2 As String) As System.Data.DataTable
        If Datatable1.Columns.Contains(Column1) = True And Datatable1.Columns.Contains(Column2) = True Then
            Dim DataView1 As New DataView(Datatable1)
            Dim eX1 As String = " ASC,"
            Dim eX2 As String = " DESC"

            DataView1.Sort = Column1 & SPACE_ASC_DESC1 & Column2 & SPACE_ASC_DESC2
            'MsgBox(DataView1.Count)

            Dim Data_table_temp As New System.Data.DataTable
            Data_table_temp = Datatable1.Clone
            Data_table_temp.Rows.Clear()

            Dim k As Integer = 0
            For i = 0 To DataView1.Count - 1
                Dim Data_row1 As DataRow
                Data_row1 = DataView1(i).Row
                Data_table_temp.Rows.Add()
                For j = 0 To Datatable1.Columns.Count - 1
                    Data_table_temp.Rows(k).Item(j) = Data_row1.Item(j)
                Next
                k = k + 1
            Next
            Return Data_table_temp
        End If




    End Function


    Public Function Incarca_existing_textstyles_to_combobox(ByVal Combo_textstyles As Windows.Forms.ComboBox)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument



                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction



                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    If Combo_textstyles.Items.Count > 0 Then Combo_textstyles.Items.Clear()


                    For Each Text_id As ObjectId In Text_style_table
                        Dim TextStyle1 As TextStyleTableRecord = Trans1.GetObject(Text_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Combo_textstyles.Items.Add(TextStyle1.Name)
                    Next




                End Using ' asta e de la trans1

            End Using

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function

    Public Function GET_distanta_Double_XY(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
        Return ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
    End Function

    Public Function GET_distanta3d_Double_with_CSF(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal CSF As Double) As Double
        Dim hd_dist_grid As Double = ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
        Dim hd_dist_ground = hd_dist_grid / CSF
        Return (hd_dist_ground ^ 2 + (z1 - z2) ^ 2) ^ 0.5
    End Function

    Public Function Incarca_existing_LINETYPES_to_combobox(ByVal Combo_LINETYPES As Windows.Forms.ComboBox)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument



                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    Dim LINETYPE_table As Autodesk.AutoCAD.DatabaseServices.LinetypeTable = Trans1.GetObject(ThisDrawing.Database.LinetypeTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Combo_LINETYPES.Items.Clear()


                    For Each linetype_id As ObjectId In LINETYPE_table
                        Dim LineType As LinetypeTableRecord = Trans1.GetObject(linetype_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Combo_LINETYPES.Items.Add(LineType.Name)
                    Next




                End Using ' asta e de la trans1

            End Using
            Combo_LINETYPES.SelectedIndex = 0

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try

    End Function
    Public Function Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(ByVal NPS_inches As Double) As Double
        Select Case NPS_inches
            Case 0.125 '1/2
                Return 5.15
            Case 0.25 '1/4
                Return 6.85
            Case 0.375 '3/8
                Return 8.55
            Case 0.5 '1/2
                Return 10.65
            Case 0.75 '3/4
                Return 13.35
            Case 1
                Return 16.7
            Case 1.25
                Return 21.1
            Case 1.5
                Return 24.15
            Case 2
                Return 30.15
            Case 2.5
                Return 36.5
            Case 3
                Return 44.45
            Case 3.5
                Return 50.8
            Case 4
                Return 57.15
            Case 5
                Return 70.65
            Case 6
                Return 84.15
            Case 8
                Return 109.55
            Case 10
                Return 136.55
            Case 12
                Return 161.95
            Case 14
                Return 177.8
            Case 16
                Return 203.2
            Case 18
                Return 228.5
            Case 20
                Return 254
            Case 22
                Return 279.5
            Case 24
                Return 304.8
            Case 26
                Return 330
            Case 28
                Return 355.5
            Case 30
                Return 381
            Case 32
                Return 406.5
            Case 34
                Return 432
            Case 36
                Return 457.2
            Case 38
                Return 482.5
            Case 40
                Return 508
            Case 42
                Return 533.5
            Case 44
                Return 559
            Case 46
                Return 584
            Case 48
                Return 609.5
            Case 50
                Return 635
            Case 52
                Return 660.5
            Case 54
                Return 686
            Case 56
                Return 711
            Case 58
                Return 736.5
            Case 60
                Return 762.0
            Case 62
                Return 787.5
            Case 64
                Return 813
            Case 66
                Return 838.0
            Case 68
                Return 863.5
            Case 70
                Return 889
            Case 72
                Return 914.5
            Case 74
                Return 940
            Case 76
                Return 965
            Case 78
                Return 990.5
            Case 80
                Return 1016
            Case Else
                Return 203.2
        End Select
    End Function
    Public Function Cel_mai_aproape_multiplu(ByVal Numar As Single, ByVal Multiplu As Double) As Single

        Dim Numar_cautat As Single = Numar / Multiplu

        If Numar_cautat - Int(Numar_cautat) < 0.5 Then
            Return Int(Numar_cautat) * Multiplu
        Else
            Return Int(Numar_cautat + 1) * Multiplu
        End If

    End Function

    Public Function Creaza_layer_cu_linetype_si_lineweight(ByVal Layername1 As String, ByVal Culoare As Integer, ByVal Linetype_name As String, ByVal Lineweight1 As Autodesk.AutoCAD.DatabaseServices.LineWeight, ByVal Descriptie As String, ByVal Plot As Boolean, ByVal Overwrite_layer As Boolean)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using lock As DocumentLock = ThisDrawing.LockDocument
                Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database = ThisDrawing.Database
                    Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable
                    LayerTable1 = Trans2.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim Linetype_table As Autodesk.AutoCAD.DatabaseServices.LinetypeTable = Trans2.GetObject(ThisDrawing.Database.LinetypeTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    If Overwrite_layer = True Then
                        If LayerTable1.Has(Layername1) = True Then
                            LayerTable1.UpgradeOpen()
                            Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                            LayerTableRecord1 = Trans2.GetObject(LayerTable1(Layername1), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            LayerTableRecord1.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare)
                            LayerTableRecord1.IsPlottable = Plot
                            LayerTableRecord1.LineWeight = Lineweight1
                            If Linetype_table.Has(Linetype_name) = True Then
                                LayerTableRecord1.LinetypeObjectId = Linetype_table.Item(Linetype_name)
                            Else

                                Dim Fisier_Creat As Boolean = False

                                Select Case Linetype_name.ToUpper
                                    Case "TCCENTER"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCCENTER", Fisier, True)
                                    Case "TCDASH"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"

                                        Incarca_linetype_Cu_specificare_fisier("TCDASH", Fisier, True)

                                    Case "TC_FENCE"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_FENCE", Fisier, True)

                                    Case "TCDOT4"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCDOT4", Fisier, True)

                                    Case "TCPHANTOM4"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCPHANTOM4", Fisier, True)

                                    Case "TCHIDDEN"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCHIDDEN", Fisier, True)

                                    Case "TCHIDDEN2"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCHIDDEN2", Fisier, True)

                                    Case "TCDASHED"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCDASHED", Fisier, True)

                                    Case "TCPHANTOM2"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCPHANTOM2", Fisier, True)

                                    Case "TCPHANTOMX2"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCPHANTOMX2", Fisier, True)

                                    Case "TCCENTER3"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCCENTER3", Fisier, True)

                                    Case "TC_FAC_FENCE"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_FAC_FENCE", Fisier, True)

                                    Case "TC_UG_TEL"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_UG_TEL", Fisier, True)

                                    Case "TC_FOREIGN_PIPE"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_FOREIGN_PIPE", Fisier, True)

                                    Case "TC_GAS_LINE"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_GAS_LINE", Fisier, True)

                                    Case "TC_OIL_LINE"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_OIL_LINE", Fisier, True)

                                    Case "TC_TRACK_S"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_TRACK_S", Fisier, True)

                                    Case "TC_TRACK_A"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_TRACK_A", Fisier, True)

                                    Case "TC_TRACK_M"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_TRACK_M", Fisier, True)

                                    Case "TCDASH3"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCDASH3", Fisier, True)

                                    Case "TC_TELEPHONE"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TC_TELEPHONE", Fisier, True)

                                    Case "PHANTOM2"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("PHANTOM2", Fisier, True)

                                    Case "TCDASH2"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCDASH2", Fisier, True)

                                    Case "TCDOT2"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"
                                        Incarca_linetype_Cu_specificare_fisier("TCDOT2", Fisier, True)
                                    Case "DOT2"
                                        Dim Locatie_fisiere_Blocuri = Locatie.Locatie_blocuri_my_docs_
                                        If Fisier_Creat = False Then
                                            Creaza_transcanada_linetype_file(Locatie_fisiere_Blocuri, "TCPL.lin")
                                            Fisier_Creat = True
                                        End If
                                        Dim Fisier As String = Locatie_fisiere_Blocuri & "\TCPL.lin"

                                        Incarca_linetype_Cu_specificare_fisier("DOT2", Fisier, True)

                                End Select


                            End If

                            LayerTableRecord1.Description = Descriptie
                            Trans2.Commit()



                        End If
                    End If



                    If LayerTable1.Has(Layername1) = False Then
                        LayerTable1.UpgradeOpen()
                        Dim LayerTableRecord1 As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord = New Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                        LayerTableRecord1.Name = Layername1
                        LayerTableRecord1.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare)
                        LayerTableRecord1.IsPlottable = Plot
                        LayerTableRecord1.LineWeight = Lineweight1
                        If Linetype_table.Has(Linetype_name) = True Then
                            LayerTableRecord1.LinetypeObjectId = Linetype_table.Item(Linetype_name)
                        End If

                        LayerTable1.Add(LayerTableRecord1)

                        Trans2.AddNewlyCreatedDBObject(LayerTableRecord1, True)

                        LayerTableRecord1.Description = Descriptie
                        Trans2.Commit()



                    End If



                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Function Creaza_transcanada_linetype_file(ByVal Locatie1 As String, ByVal Nume_fisier As String)
        Try

            If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Locatie1) = False Then
                Microsoft.VisualBasic.FileIO.FileSystem.CreateDirectory(Locatie1)
            End If
            Dim Fisier As String = Locatie1 & "\" & Nume_fisier

            Dim fs As IO.FileStream = System.IO.File.Create(Fisier)
            fs.Close()

            Dim StreamWriter1 As New System.IO.StreamWriter(Fisier)
            Using StreamWriter1
                StreamWriter1.Write(
                "*{ Wide Dash },{ Wide Dash }" & vbCrLf &
                "A, 4, -2, 4, -2" & vbCrLf &
                "*{ -E- },Dash triple-dot" & vbCrLf &
                "A,4,-0.75,[" & Chr(34) & "E" & Chr(34) & ",ROMANS,s=0.15,r=0,x=-0.05,y=-0.07],-0.75" & vbCrLf &
                "*TCCENTER,--- - --- - --- - --- - --- - --- - --- - --- -" & vbCrLf &
                "A,20,-1,2,-1" & vbCrLf &
                "*TCCENTER2,___ _ ___ _ ___ _ ___ _ ___ _ ___ _ ___ _ ___ _" & vbCrLf &
                "A,20,-3,3,3" & vbCrLf &
                "*TCCENTER3,____ _ ____ _ ____ _ ____ _ ____ _ ____ _ ____ " & vbCrLf &
                "A, 32, -6, 6, -6" & vbCrLf &
                "*TCDOT2,..............................................." & vbCrLf &
                "A, 0.0, -3.0" & vbCrLf &
                "*DOT2,Dot (.5x) ....................................." & vbCrLf &
                "A, 0, -3.175" & vbCrLf &
                "*TCDOT4,. . . . . . . ." & vbCrLf &
                "A,0,-0.625" & vbCrLf &
                "*DOT4,. . . . ." & vbCrLf &
                "A,0,-1" & vbCrLf &
                "*TCDASH,-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --" & vbCrLf &
                "A,1,-1" & vbCrLf &
                "*TCDASH2,- - - - - - - - - - - - - - - - - - - - - - - -" & vbCrLf &
                "A, 2, -2" & vbCrLf &
                "*TCDASH3,----------  ----------  ----------  ---------- " & vbCrLf &
                "A, 10, -2" & vbCrLf &
                "*TCDASHED,__ __ __ __ __ __ __ __ __ __ __ __ __ __ __ __" & vbCrLf &
                "A,12,-6" & vbCrLf &
                "*TCHIDDEN,- - - - - - - - - - - - - - - - - - - - - - - -" & vbCrLf &
                "A,1,-1" & vbCrLf &
                "*TCHIDDEN2,_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _" & vbCrLf &
                "A,3,-2" & vbCrLf &
                "*HIDDEN2,_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _" & vbCrLf &
                "A, 3, -2" & vbCrLf &
                "*HIDDEN,- - - - - - - - - - - - - - - - - - - - - - - -" & vbCrLf &
                "A, 1, -1" & vbCrLf &
                "*TCPHANTOM,--------- -- --------- -- --------- -- --------" & vbCrLf &
                "A, 20, -1, 2, -1, 2, -1" & vbCrLf &
                "*TCPHANTOM2,___ _ _ ___ _ _ ___ _ _ ___ _ _ ___ _ _ ___ _ _" & vbCrLf &
                "A, 16, -3, 3, -3, 3, -3" & vbCrLf &
                "*TCPHANTOMX2,____________    ____    ____    ____________" & vbCrLf &
                "A, 64, -12, 12, -12, 12, -12" & vbCrLf &
                "*PHANTOM2,Phantom (.5x) ___ _ _ ___ _ _ ___ _ _ ___ _ _" & vbCrLf &
                "A, 0.625, -0.125, 0.125, -0.125, 0.125, -0.125" & vbCrLf &
                "*TCPHANTOM4," & vbCrLf &
                "A,5,-0.75,1.5,-0.75,1.5,-0.75" & vbCrLf &
                "*PHANTOM4," & vbCrLf &
                "A,5,-0.75,1.5,-0.75,1.5,-0.75" & vbCrLf &
                "*TC_POWER,--P----P----P----P----P----P--" & vbCrLf &
                "A,18,-1.75,[" & Chr(34) & "P" & Chr(34) & ",ROMANS,s=1.75,r=0,x=-0.85,y=-0.85],-1.75" & vbCrLf &
                "*TC_BURIED_CABLE,--C----C----C----C----C----C--" & vbCrLf &
                "A,18,-1.75,[" & Chr(34) & "C" & Chr(34) & ",ROMANS,s=1.75,r=0,x=-0.85,y=-0.85],-1.75" & vbCrLf &
                "*TC_FENCE,--x----x----x----x----x----x--" & vbCrLf &
                "A,18,-1.25,[" & Chr(34) & "x" & Chr(34) & ",ROMANS,s=3,r=0,x=-1,y=-1],-2" & vbCrLf &
                "*TC_FAC_FENCE,----//----//----//----//----" & vbCrLf &
                "A,13,-1.5,[TRACK1,ltypeshp.shx,s=1.5,r=315,x=0,y=0],0,-1.25,[TRACK1,ltypeshp.shx,s=1.5,r=315,x=0,y=0],-1.5" & vbCrLf &
                "*TC_UG_TEL,----0-----0----" & vbCrLf &
                "A,13,[CIRC1,ltypeshp.shx,s=1.5,r=0,x=0,y=0],-3,[" & Chr(34) & "T" & Chr(34) & ",ROMANS,s=1.75,r=0,x=-2.15,y=-1],5" & vbCrLf &
                "*TC_TELEPHONE,--T----T----T----T----T----T--" & vbCrLf &
                "A,18,-1.75,[" & Chr(34) & "T" & Chr(34) & ",ROMANS,s=1.75,r=0,x=-0.85,y=-0.85],-1.5" & vbCrLf &
                "*TC_FOREIGN_PIPE,--| |----| |----| |--" & vbCrLf &
                "A,3,[TRACK1,ltypeshp.shx,s=0.5,r=0,x=0,y=0],-1.5,[TRACK1,ltypeshp.shx,s=0.5,r=0,x=0,y=0],3" & vbCrLf &
                "*TC_GAS_LINE,----GAS----GAS----GAS----GAS" & vbCrLf &
                "A,4,-1,[" & Chr(34) & "GAS" & Chr(34) & ",ROMANS,s=1.25,r=0,x=-0.5,y=-0.625],-3.5" & vbCrLf &
                "*TC_OIL_LINE,----OIL----OIL----OIL----OIL" & vbCrLf &
                "A,4,-1,[" & Chr(34) & "OIL" & Chr(34) & ",ROMANS,s=1.25,r=0,x=-0.5,y=-0.625],-3" & vbCrLf &
                "*TC_TRACK_S,----|----|----|----|----|----" & vbCrLf &
                "A,3,[TRACK1,ltypeshp.shx,s=1,r=0,x=0,y=0],3" & vbCrLf &
                "*TC_TRACK_A,--|--  --|--  --|--  --|--" & vbCrLf &
                "A,3,[TRACK1,ltypeshp.shx,s=0.5,r=0,x=0,y=0],3,-3" & vbCrLf &
                "*TC_TRACK_M,----||----||----||----||----" & vbCrLf &
                "A,3,[TRACK1,ltypeshp.shx,s=0.5,r=0,x=0,y=0],0.5,[TRACK1,ltypeshp.shx,s=0.5,r=0,x=0,y=0],3" & vbCrLf &
                "*SEISMIC,Dash dot __ . __ . __ . __ . __ . __ . __ . __" & vbCrLf &
                "A, 0.5, -0.25, 0, -0.25"
                )
                StreamWriter1.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
            SET_FILEDIA_TO_1()
            Application.SetSystemVariable("LTSCALE", 1)
        End Try
    End Function


    Public Function extrage_chainage_din_text_de_la_inceputul_textului(ByVal string1 As String) As String
        Try
            Dim Chainage As String = ""
            If string1.Contains("+") = True Then
                For i = 1 To string1.Length
                    Dim Litera As String = Mid(string1, i, 1)

                    Select Case Litera
                        Case "."
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "0"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "1"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "2"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "3"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "4"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "5"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "6"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "7"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "8"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "9"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "+"
                            If i - 1 = Len(Chainage) Then Chainage = Chainage & Litera
                        Case "-"
                            If i = 1 Then Chainage = Chainage & Litera
                        Case Else
                            Exit For
                    End Select
                Next
            End If


            Return Chainage

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function extrage_chainage_din_text_de_la_sfarsitul_textului(ByVal string1 As String) As String
        Try
            Dim Chainage As String = ""
            If string1.Contains("+") = True Then
                For i = string1.Length To 1 Step -1
                    Dim Litera As String = Mid(string1, i, 1)

                    Select Case Litera
                        Case "."
                            Chainage = Litera & Chainage
                        Case "0"
                            Chainage = Litera & Chainage
                        Case "1"
                            Chainage = Litera & Chainage
                        Case "2"
                            Chainage = Litera & Chainage
                        Case "3"
                            Chainage = Litera & Chainage
                        Case "4"
                            Chainage = Litera & Chainage
                        Case "5"
                            Chainage = Litera & Chainage
                        Case "6"
                            Chainage = Litera & Chainage
                        Case "7"
                            Chainage = Litera & Chainage
                        Case "8"
                            Chainage = Litera & Chainage
                        Case "9"
                            Chainage = Litera & Chainage
                        Case "+"
                            Chainage = Litera & Chainage
                        Case "-"
                            If i = 1 Then Chainage = Litera & Chainage
                        Case Else
                            Exit For
                    End Select
                Next
            End If


            Return Chainage

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Incarca_linetype_Cu_specificare_fisier(ByVal LineType_name As String, ByVal Linetype_file As String, ByVal overwrite_from_disk As Boolean)
        Try
            Application.SetSystemVariable("FILEDIA", 0)
            If IO.File.Exists(Linetype_file) = True Then
                Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

                Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim Linetype_table As Autodesk.AutoCAD.DatabaseServices.LinetypeTable = Trans1.GetObject(ThisDrawing.Database.LinetypeTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        Dim Linetype_table_record As LinetypeTableRecord
                        If Linetype_table.Has(LineType_name) = False Then
                            Linetype_table.UpgradeOpen()
                            ThisDrawing.Database.LoadLineTypeFile(LineType_name, Linetype_file)
                        Else
                            If overwrite_from_disk = True Then


                                ThisDrawing.SendStringToExecute("-linetype Load " & LineType_name & vbCr & Linetype_file & vbCr & "Yes" & vbCrLf, False, False, True)

                                'Dim AcadDocObj As Object = Application.AcadApplication.GetType.InvokeMember("ActiveDocument", Reflection.BindingFlags.GetProperty, Nothing, Application.AcadApplication, Nothing)
                                'Dim DataArray(0) As Object
                                ' DataArray(0) = "-linetype Load " & LineType_name & vbCr & Linetype_file & vbCr & "Yes" & vbCrLf
                                'AcadDocObj.GetType.InvokeMember("SendCommand", Reflection.BindingFlags.InvokeMethod, Nothing, AcadDocObj, DataArray)


                            End If
                        End If
                        Trans1.Commit()
                    End Using ' asta e de la trans1
                End Using
            End If
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Catch ex As Exception

            SET_FILEDIA_TO_1()
            MsgBox(ex.Message)
        End Try

    End Function

    Public Sub SET_FILEDIA_TO_1()

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("FILEDIA" & vbCr & "1" & vbCr, False, False, True)
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("MSLTSCALE" & vbCr & "0" & vbCr, False, False, True)
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("PSLTSCALE" & vbCr & "1" & vbCr, False, False, True)
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("LTSCALE" & vbCr & "1" & vbCr, False, False, True)

        ' Dim AcadDocObj As Object = Application.AcadApplication.GetType.InvokeMember("ActiveDocument", Reflection.BindingFlags.GetProperty, Nothing, Application.AcadApplication, Nothing)
        'Dim DataArray(0) As Object
        'DataArray(0) = "FILEDIA" & vbCr & "1 "
        'AcadDocObj.GetType.InvokeMember("SendCommand", Reflection.BindingFlags.InvokeMethod, Nothing, AcadDocObj, DataArray)
    End Sub
    Public Function Creaza_group_of_dbobjects(ByVal ObjectId1 As ObjectIdCollection, ByVal Nume_grup As String) As Group
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If ObjectId1.Count > 0 Then
                Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim GroupDictionary As DBDictionary = Trans1.GetObject(ThisDrawing.Database.GroupDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Dim Increment As Double = 0
123:
                        If GroupDictionary.Contains(Nume_grup) = False Then
                            Dim Grup1 As New Group(Nume_grup, True)
                            Grup1.SetColorIndex(0)
                            GroupDictionary.SetAt(Nume_grup, Grup1)
                            Trans1.AddNewlyCreatedDBObject(Grup1, True)
                            Grup1.InsertAt(0, ObjectId1)
                        Else
                            Nume_grup = Nume_grup & Increment
                            Increment = Increment + 1
                            GoTo 123
                        End If
                        Trans1.Commit()
                    End Using ' asta e de la trans1
                End Using
            End If
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function delete_DBnull_rows_from_data_table(ByVal Datatable1 As System.Data.DataTable, ByVal Column1 As String) As System.Data.DataTable
        If Datatable1.Columns.Contains(Column1) = True Then
            Dim nr As Double = Datatable1.Rows.Count
            For i = 0 To nr - 1
                If i < Datatable1.Rows.Count Then
                    If IsDBNull(Datatable1.Rows(i).Item(Column1)) = True Then
                        Datatable1.Rows.RemoveAt(i)
                        If i > 0 Then i = i - 1
                    End If
                End If
            Next
            Return Datatable1
        End If
    End Function

    Public Function extrage_numar_din_text_de_la_sfarsitul_textului(ByVal string1 As String) As String
        Try
            Dim Numar As String = ""

            For i = string1.Length To 1 Step -1
                Dim Litera As String = Mid(string1, i, 1)

                Select Case Litera
                    Case "."
                        Numar = Litera & Numar
                    Case "0"
                        Numar = Litera & Numar
                    Case "1"
                        Numar = Litera & Numar
                    Case "2"
                        Numar = Litera & Numar
                    Case "3"
                        Numar = Litera & Numar
                    Case "4"
                        Numar = Litera & Numar
                    Case "5"
                        Numar = Litera & Numar
                    Case "6"
                        Numar = Litera & Numar
                    Case "7"
                        Numar = Litera & Numar
                    Case "8"
                        Numar = Litera & Numar
                    Case "9"
                        Numar = Litera & Numar
                    Case "-"
                        If i = 1 Then Numar = Litera & Numar
                    Case Else
                        Exit For
                End Select
            Next



            Return Numar

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function extrage_numar_din_text(ByVal string1 As String) As String
        Try
            Dim Numar As String = ""

            For i = 1 To string1.Length
                Dim Litera As String = Mid(string1, i, 1)

                Select Case Litera
                    Case "."
                        Numar = Numar & Litera
                    Case "0"
                        Numar = Numar & Litera
                    Case "1"
                        Numar = Numar & Litera
                    Case "2"
                        Numar = Numar & Litera
                    Case "3"
                        Numar = Numar & Litera
                    Case "4"
                        Numar = Numar & Litera
                    Case "5"
                        Numar = Numar & Litera
                    Case "6"
                        Numar = Numar & Litera
                    Case "7"
                        Numar = Numar & Litera
                    Case "8"
                        Numar = Numar & Litera
                    Case "9"
                        Numar = Numar & Litera
                    Case "-"
                        If i = 1 Then Numar = Numar & Litera
                End Select
            Next



            Return Numar

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function extrage_numar_din_text_de_la_inceputul_textului(ByVal string1 As String) As String
        Try
            Dim Numar As String = ""

            For i = 1 To string1.Length
                Dim Litera As String = Mid(string1, i, 1)

                Select Case Litera
                    Case "."
                        Numar = Numar & Litera
                    Case "0"
                        Numar = Numar & Litera
                    Case "1"
                        Numar = Numar & Litera
                    Case "2"
                        Numar = Numar & Litera
                    Case "3"
                        Numar = Numar & Litera
                    Case "4"
                        Numar = Numar & Litera
                    Case "5"
                        Numar = Numar & Litera
                    Case "6"
                        Numar = Numar & Litera
                    Case "7"
                        Numar = Numar & Litera
                    Case "8"
                        Numar = Numar & Litera
                    Case "9"
                        Numar = Numar & Litera
                    Case "-"
                        If i = 1 Then Numar = Numar & Litera
                    Case Else
                        Exit For
                End Select
            Next



            Return Numar

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function extrage_STATION_din_text_de_la_inceputul_textului(ByVal string1 As String) As String
        Try
            Dim Numar As String = ""

            For i = 1 To string1.Length
                Dim Litera As String = Mid(string1, i, 1)

                Select Case Litera
                    Case "0"
                        Numar = Numar & Litera
                    Case "1"
                        Numar = Numar & Litera
                    Case "2"
                        Numar = Numar & Litera
                    Case "3"
                        Numar = Numar & Litera
                    Case "4"
                        Numar = Numar & Litera
                    Case "5"
                        Numar = Numar & Litera
                    Case "6"
                        Numar = Numar & Litera
                    Case "7"
                        Numar = Numar & Litera
                    Case "8"
                        Numar = Numar & Litera
                    Case "9"
                        Numar = Numar & Litera
                    Case "+"
                        Numar = Numar & Litera
                    Case Else
                        Exit For
                End Select
            Next



            Return Numar

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function calculeaza_chainage_for_REROUTE(ByRef Curba_noua As Curve, ByRef Curba_veche As Curve, ByVal Point_on_poly As Point3d, ByVal Point_zero_old As Point3d) As Double

        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try



            Dim Chainage_at_common_point_old As Double
            Dim Chainage_at_common_point_new As Double
            Dim Diferenta_chainage As Double

            Dim Poly1 As Polyline
            Dim Poly3D As Polyline3d

            Dim Poly2 As Polyline
            Dim Poly3D2 As Polyline3d




            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                Dim Ent1 As Entity
                Ent1 = Curba_veche.ObjectId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                Dim Ent2 As Entity
                Ent2 = Curba_noua.ObjectId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                    Poly1 = Ent1
                    Poly2 = Ent2



                    Dim Point_zero_new As New Point3d


                    Chainage_at_common_point_old = Poly1.GetDistAtPoint(Point_zero_old)
                    Point_zero_new = Poly2.GetClosestPointTo(Point_zero_old, Vector3d.ZAxis, False)
                    Chainage_at_common_point_new = Poly2.GetDistAtPoint(Point_zero_new)
                    Diferenta_chainage = Chainage_at_common_point_new - Chainage_at_common_point_old



                ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                    Poly3D = Ent1
                    Poly3D2 = Ent2



                    Dim Point_zero_new As New Point3d


                    Chainage_at_common_point_old = Poly3D.GetDistAtPoint(Point_zero_old)
                    Point_zero_new = Poly3D2.GetClosestPointTo(Point_zero_old, Vector3d.ZAxis, False)
                    Chainage_at_common_point_new = Poly3D2.GetDistAtPoint(Point_zero_new)
                    Diferenta_chainage = Chainage_at_common_point_new - Chainage_at_common_point_old



                Else
                    Editor1.WriteMessage("No Polylines")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Function
                End If


                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly2) = False Then
                    Distanta_pana_la_xing = Poly2.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D2) = False Then

                    Distanta_pana_la_xing = Poly3D2.GetDistAtPoint(Point_on_poly)
                End If


                Return Distanta_pana_la_xing - Diferenta_chainage


            End Using



        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Function

    Public Function Creaza_Dim_style_AND_text_style_ROMANS()
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)




                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Dim_style_table As Autodesk.AutoCAD.DatabaseServices.DimStyleTable = Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Leader_style_table As DBDictionary = Trans1.GetObject(ThisDrawing.Database.MLeaderStyleDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Dim_style_universal As DimStyleTableRecord
                    Dim Dim_style_universal_min_cvr As DimStyleTableRecord
                    Dim Text_style_Text25 As TextStyleTableRecord
                    Dim Text_style_Text40 As TextStyleTableRecord

                    For Each TextStyle_id As ObjectId In Text_style_table
                        Dim TextStyle As TextStyleTableRecord = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If TextStyle.Name.ToUpper = "ROMANS" Then
                            Text_style_Text25 = TextStyle
                            Text_style_Text25.UpgradeOpen()
                            Text_style_Text25.TextSize = 0
                            Text_style_Text25.ObliquingAngle = 0
                            Text_style_Text25.FileName = "romans.shx"
                            Text_style_Text25.XScale = 1.0
                        End If
                    Next

                    If IsNothing(Text_style_Text25) = True Then
                        Text_style_Text25 = New TextStyleTableRecord
                        Text_style_Text25.Name = "ROMANS"

                        Text_style_Text25.TextSize = 0
                        Text_style_Text25.ObliquingAngle = 0
                        Text_style_Text25.FileName = "romans.shx"
                        Text_style_Text25.XScale = 1.0
                        Text_style_table.Add(Text_style_Text25)
                        Trans1.AddNewlyCreatedDBObject(Text_style_Text25, True)

                    End If


                    Application.SetSystemVariable("TEXTSTYLE", "ROMANS")






                    Dim Arrowid As ObjectId
                    Dim Arrowid_OPEN30 As ObjectId
                    Dim BlockTable As BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                    Application.SetSystemVariable("DIMBLK", "_Open30")
                    If BlockTable.Has("_Open30") = True Then
                        Arrowid_OPEN30 = BlockTable("_Open30")
                    End If


                    For Each dimStyle_id As ObjectId In Dim_style_table
                        Dim dimStyle As DimStyleTableRecord = Trans1.GetObject(dimStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If dimStyle.Name.ToUpper = "HMM" Then
                            Dim_style_universal = dimStyle
                            Dim_style_universal.UpgradeOpen()
                            With Dim_style_universal

                                .Dimadec = 0
                                .Dimalt = False
                                .Dimaltd = 2
                                .Dimaltf = 25.4
                                .Dimaltrnd = 0
                                .Dimalttd = 2
                                .Dimalttz = 0
                                .Dimaltu = 2
                                .Dimaltz = 0
                                .Dimapost = ""
                                .Dimarcsym = 0
                                .Dimasz = 3
                                .Dimatfit = 0
                                .Dimaunit = 0
                                .Dimazin = 0
                                .Dimcen = 0.05
                                .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                .Dimdec = 1
                                .Dimdle = 0
                                .Dimdli = 0.3785
                                .Dimdsep = ".c"
                                .Dimexe = 1
                                .Dimexo = 1
                                .Dimfrac = 0
                                .Dimfxlen = 1
                                .DimfxlenOn = False
                                .Dimgap = 0.9
                                .Dimjogang = 0.785398163397448
                                .Dimjust = 0
                                .Dimblk = Arrowid_OPEN30
                                .Dimldrblk = Arrowid
                                .Dimlfac = 1
                                .Dimlim = False

                                .Dimlunit = 2
                                .Dimlwd = LineWeight.ByBlock
                                .Dimlwe = LineWeight.ByBlock

                                .Dimpost = ""
                                .Dimrnd = 0
                                .Dimsah = False
                                .Dimscale = 1
                                .Dimsd1 = False
                                .Dimsd2 = False
                                .Dimse1 = False
                                .Dimse2 = False
                                .Dimsoxd = False
                                .Dimtad = 1
                                .Dimtdec = 1
                                .Dimtfac = 1
                                .Dimtfill = 0
                                .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)
                                .Dimtih = False
                                .Dimtix = False
                                .Dimtm = 0
                                .Dimtmove = 0
                                .Dimtofl = False
                                .Dimtoh = False
                                .Dimtol = False
                                .Dimtolj = 1
                                .Dimtp = 0
                                .Dimtsz = 0
                                .Dimtvp = 0
                                .Dimtxsty = Text_style_Text25.ObjectId
                                .Dimtxt = 2.5
                                .Dimtxtdirection = False
                                .Dimtzin = 0
                                .Dimupt = False
                                .Dimzin = 0


                            End With

                        End If

                        If dimStyle.Name.ToUpper = "HMM MIN CVR" Then
                            Dim_style_universal_min_cvr = dimStyle
                            Dim_style_universal_min_cvr.UpgradeOpen()
                            With Dim_style_universal_min_cvr

                                .Dimadec = 1
                                .Dimalt = False
                                .Dimaltd = 2
                                .Dimaltf = 25.4
                                .Dimaltrnd = 0
                                .Dimalttd = 2
                                .Dimalttz = 0
                                .Dimaltu = 2
                                .Dimaltz = 0
                                .Dimapost = ""
                                .Dimarcsym = 0
                                .Dimasz = 3
                                .Dimatfit = 0
                                .Dimaunit = 0
                                .Dimazin = 0
                                .Dimblk = Arrowid_OPEN30
                                .Dimblk1 = Arrowid
                                .Dimblk2 = Arrowid
                                .Dimcen = 1.5
                                .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                .Dimdec = 1
                                .Dimdle = 0
                                .Dimdli = 0
                                .Dimdsep = ".c"
                                .Dimexe = 2
                                .Dimexo = 2
                                .Dimfrac = 0
                                .Dimfxlen = 1
                                .DimfxlenOn = False
                                .Dimgap = 0.65
                                .Dimjogang = 0.785398163397448
                                .Dimjust = 0
                                .Dimldrblk = Arrowid
                                .Dimlfac = 1
                                .Dimlim = False
                                .Dimltex1 = ThisDrawing.Database.ByBlockLinetype
                                .Dimltex2 = ThisDrawing.Database.ByBlockLinetype
                                .Dimltype = ThisDrawing.Database.ByBlockLinetype
                                .Dimlunit = 2
                                .Dimlwd = LineWeight.ByBlock
                                .Dimlwe = LineWeight.ByBlock
                                .Dimpost = ""
                                .Dimrnd = 0
                                .Dimsah = False
                                .Dimscale = 1
                                .Dimsd1 = False
                                .Dimsd2 = False
                                .Dimse1 = True
                                .Dimse2 = True
                                .Dimsoxd = False
                                .Dimtad = 0
                                .Dimtdec = 1
                                .Dimtfac = 1
                                .Dimtfill = 0
                                .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)
                                .Dimtih = True
                                .Dimtix = False
                                .Dimtm = 0
                                .Dimtmove = 0
                                .Dimtofl = False
                                .Dimtoh = True
                                .Dimtol = False
                                .Dimtolj = 1
                                .Dimtp = 0
                                .Dimtsz = 0
                                .Dimtvp = 0
                                .Dimtxsty = Text_style_Text25.ObjectId
                                .Dimtxt = 2.5
                                .Dimtxtdirection = False
                                .Dimtzin = 0
                                .Dimupt = False
                                .Dimzin = 0



                            End With

                        End If
                    Next

                    If IsNothing(Dim_style_universal) = True Then
                        Dim_style_universal = New DimStyleTableRecord
                        Dim_style_universal.Name = "HMM"
                        With Dim_style_universal
                            .Dimadec = 0
                            .Dimalt = False
                            .Dimaltd = 2
                            .Dimaltf = 25.4
                            .Dimaltrnd = 0
                            .Dimalttd = 2
                            .Dimalttz = 0
                            .Dimaltu = 2
                            .Dimaltz = 0
                            .Dimapost = ""
                            .Dimarcsym = 0
                            .Dimasz = 3
                            .Dimatfit = 0
                            .Dimaunit = 0
                            .Dimazin = 0
                            .Dimcen = 0.05
                            .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                            .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                            .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                            .Dimdec = 1
                            .Dimdle = 0
                            .Dimdli = 0.3785
                            .Dimdsep = ".c"
                            .Dimexe = 1
                            .Dimexo = 1
                            .Dimfrac = 0
                            .Dimfxlen = 1
                            .DimfxlenOn = False
                            .Dimgap = 0.9
                            .Dimjogang = 0.785398163397448
                            .Dimjust = 0
                            .Dimblk = Arrowid_OPEN30
                            .Dimldrblk = Arrowid
                            .Dimlfac = 1
                            .Dimlim = False

                            .Dimlunit = 2
                            .Dimlwd = LineWeight.ByBlock
                            .Dimlwe = LineWeight.ByBlock

                            .Dimpost = ""
                            .Dimrnd = 0
                            .Dimsah = False
                            .Dimscale = 1
                            .Dimsd1 = False
                            .Dimsd2 = False
                            .Dimse1 = False
                            .Dimse2 = False
                            .Dimsoxd = False
                            .Dimtad = 1
                            .Dimtdec = 1
                            .Dimtfac = 1
                            .Dimtfill = 0
                            .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)
                            .Dimtih = False
                            .Dimtix = False
                            .Dimtm = 0
                            .Dimtmove = 0
                            .Dimtofl = False
                            .Dimtoh = False
                            .Dimtol = False
                            .Dimtolj = 1
                            .Dimtp = 0
                            .Dimtsz = 0
                            .Dimtvp = 0
                            .Dimtxsty = Text_style_Text25.ObjectId
                            .Dimtxt = 2.5
                            .Dimtxtdirection = False
                            .Dimtzin = 0
                            .Dimupt = False
                            .Dimzin = 0


                        End With
                        Dim_style_table.Add(Dim_style_universal)
                        Trans1.AddNewlyCreatedDBObject(Dim_style_universal, True)

                    End If

                    If IsNothing(Dim_style_universal_min_cvr) = True Then
                        Dim_style_universal_min_cvr = New DimStyleTableRecord
                        Dim_style_universal_min_cvr.Name = "HMM MIN CVR"
                        With Dim_style_universal_min_cvr

                            .Dimadec = 1
                            .Dimalt = False
                            .Dimaltd = 2
                            .Dimaltf = 25.4
                            .Dimaltrnd = 0
                            .Dimalttd = 2
                            .Dimalttz = 0
                            .Dimaltu = 2
                            .Dimaltz = 0
                            .Dimapost = ""
                            .Dimarcsym = 0
                            .Dimasz = 3.0
                            .Dimatfit = 0
                            .Dimaunit = 0
                            .Dimazin = 0
                            .Dimblk = Arrowid_OPEN30
                            .Dimblk1 = Arrowid
                            .Dimblk2 = Arrowid
                            .Dimcen = 1.5
                            .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                            .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                            .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                            .Dimdec = 1
                            .Dimdle = 0
                            .Dimdli = 0
                            .Dimdsep = ".c"
                            .Dimexe = 2
                            .Dimexo = 2
                            .Dimfrac = 0
                            .Dimfxlen = 1
                            .DimfxlenOn = False
                            .Dimgap = 0.65
                            .Dimjogang = 0.785398163397448
                            .Dimjust = 0
                            .Dimldrblk = Arrowid
                            .Dimlfac = 1
                            .Dimlim = False
                            .Dimltex1 = ThisDrawing.Database.ByBlockLinetype
                            .Dimltex2 = ThisDrawing.Database.ByBlockLinetype
                            .Dimltype = ThisDrawing.Database.ByBlockLinetype
                            .Dimlunit = 2
                            .Dimlwd = LineWeight.ByBlock
                            .Dimlwe = LineWeight.ByBlock
                            .Dimpost = ""
                            .Dimrnd = 0
                            .Dimsah = False
                            .Dimscale = 1
                            .Dimsd1 = False
                            .Dimsd2 = False
                            .Dimse1 = True
                            .Dimse2 = True
                            .Dimsoxd = False
                            .Dimtad = 0
                            .Dimtdec = 1
                            .Dimtfac = 1
                            .Dimtfill = 0
                            .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)
                            .Dimtih = True
                            .Dimtix = False
                            .Dimtm = 0
                            .Dimtmove = 0
                            .Dimtofl = False
                            .Dimtoh = True
                            .Dimtol = False
                            .Dimtolj = 1
                            .Dimtp = 0
                            .Dimtsz = 0
                            .Dimtvp = 0
                            .Dimtxsty = Text_style_Text25.ObjectId
                            .Dimtxt = 2.5
                            .Dimtxtdirection = False
                            .Dimtzin = 0
                            .Dimupt = False
                            .Dimzin = 0

                        End With
                        Dim_style_table.Add(Dim_style_universal_min_cvr)
                        Trans1.AddNewlyCreatedDBObject(Dim_style_universal_min_cvr, True)

                    End If



                    ThisDrawing.Database.Dimstyle = Dim_style_universal.Id
                    ThisDrawing.Database.SetDimstyleData(Trans1.GetObject(Dim_style_universal.Id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead))

                    Dim Universal_leader_style As New MLeaderStyle

                    If Leader_style_table.Contains("HMM") = True Then
                        Dim ID1 As ObjectId = Leader_style_table.GetAt("HMM")
                        Universal_leader_style = Trans1.GetObject(ID1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        With Universal_leader_style
                            .Annotative = AnnotativeStates.True
                            .ArrowSize = 3.5
                            .BreakSize = 3.5
                            .DoglegLength = 3
                            .LeaderLineColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256)
                            .TextColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256)
                            .TextHeight = 2.5
                            .TextStyleId = Text_style_Text25.ObjectId
                            .ArrowSymbolId = Arrowid
                            .BlockColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 0)
                            .BlockRotation = 0
                            .BlockScale = New Autodesk.AutoCAD.Geometry.Scale3d(1, 1, 1)
                            .ContentType = ContentType.MTextContent
                            .DrawLeaderOrderType = DrawLeaderOrderType.DrawLeaderHeadFirst
                            .DrawMLeaderOrderType = DrawMLeaderOrderType.DrawLeaderFirst
                            .EnableBlockRotation = True
                            .EnableBlockScale = True
                            .EnableDogleg = True
                            .EnableFrameText = False
                            .EnableLanding = True
                            .ExtendLeaderToText = False
                            .TextAlignAlwaysLeft = True
                            .LandingGap = 2
                            .LeaderLineType = LeaderType.StraightLeader
                            .LeaderLineWeight = LineWeight.ByBlock
                            .MaxLeaderSegmentsPoints = 2
                            .Scale = 1
                            .TextAlignAlwaysLeft = False
                            .TextAlignmentType = TextAlignmentType.LeftAlignment
                            .TextAngleType = TextAngleType.HorizontalAngle
                            '.TextAttachmentDirection = TextAttachmentDirection.AttachmentHorizontal
                            '.TextAttachmentType = TextAttachmentType.AttachmentMiddleOfTop
                            'SetTextAttachmentType(Universal_leader_style, TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader)
                            'SetTextAttachmentType(Universal_leader_style, TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader)

                        End With

                    Else

                        With Universal_leader_style
                            Leader_style_table.SetAt("HMM", Universal_leader_style)
                            .Annotative = AnnotativeStates.True
                            .ArrowSize = 3.5
                            .BreakSize = 3.5
                            .DoglegLength = 3
                            .LeaderLineColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256)
                            .TextColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256)
                            .TextHeight = 2.5
                            .TextStyleId = Text_style_Text25.ObjectId
                            .ArrowSymbolId = Arrowid
                            .BlockColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 0)
                            .BlockRotation = 0
                            .BlockScale = New Autodesk.AutoCAD.Geometry.Scale3d(1, 1, 1)
                            .ContentType = ContentType.MTextContent
                            .DrawLeaderOrderType = DrawLeaderOrderType.DrawLeaderHeadFirst
                            .DrawMLeaderOrderType = DrawMLeaderOrderType.DrawLeaderFirst
                            .EnableBlockRotation = True
                            .EnableBlockScale = True
                            .EnableDogleg = True
                            .EnableFrameText = False
                            .EnableLanding = True
                            .ExtendLeaderToText = False
                            .TextAlignAlwaysLeft = True
                            .LandingGap = 2
                            .LeaderLineType = LeaderType.StraightLeader
                            .LeaderLineWeight = LineWeight.ByBlock
                            .MaxLeaderSegmentsPoints = 2
                            .Scale = 1
                            .TextAlignAlwaysLeft = False
                            .TextAlignmentType = TextAlignmentType.LeftAlignment
                            .TextAngleType = TextAngleType.HorizontalAngle
                            '.TextAttachmentDirection = TextAttachmentDirection.AttachmentHorizontal
                            '.TextAttachmentType = TextAttachmentType.AttachmentMiddleOfTop
                            'SetTextAttachmentType(Universal_leader_style, TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader)
                            'SetTextAttachmentType(Universal_leader_style, TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader)

                        End With

                        Trans1.AddNewlyCreatedDBObject(Universal_leader_style, True)

                    End If

                    ThisDrawing.Database.MLeaderstyle = Universal_leader_style.ObjectId

                    Trans1.Commit()



                End Using ' asta e de la trans1

            End Using

            Dim Locatie1 As String = Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) & "\BLOCKS"
            Creaza_transcanada_linetype_file(Locatie1, "TCPL.lin")
            Dim Fisier As String = Locatie1 & "\TCPL.lin"


            If Not Fisier = "" Then
                Incarca_linetype_Cu_specificare_fisier("{ Wide Dash }", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("{ -E- }", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCCENTER", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCCENTER2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCCENTER3", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCDOT2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("DOT2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCDOT4", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("DOT4", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCDASH", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCDASH2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCDASH3", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCDASHED", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCHIDDEN", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCHIDDEN2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("HIDDEN2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("HIDDEN", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCPHANTOM", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCPHANTOM2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCPHANTOMX2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("PHANTOM2", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TCPHANTOM4", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("PHANTOM4", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_POWER", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_BURIED_CABLE", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_FENCE", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_FAC_FENCE", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_UG_TEL", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_TELEPHONE", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_FOREIGN_PIPE", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_GAS_LINE", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_OIL_LINE", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_TRACK_S", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_TRACK_A", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("TC_TRACK_M", Fisier, True)
                Incarca_linetype_Cu_specificare_fisier("SEISMIC", Fisier, True)
            End If


            SET_FILEDIA_TO_1()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception

            MsgBox(ex.Message)
            SET_FILEDIA_TO_1()
            Application.SetSystemVariable("LTSCALE", 1)
        End Try

    End Function
    Public Function SetTextAttachmentType(ByVal mlStyle As Autodesk.AutoCAD.DatabaseServices.MLeaderStyle, ByVal textAttachmentType As TextAttachmentType, ByVal leaderDirection As LeaderDirectionType) As ErrorStatus
        'Return SetTextAttachmentType(mlStyle, textAttachmentType, leaderDirection)
    End Function
    Public Function Insereaza_block_table_record_in_drawing(ByVal Nume_fisier As String, ByVal NumeBlock As String) As BlockTableRecord


        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Editor1 = ThisDrawing.Editor
        Dim BlockTable1 As BlockTable
        Dim BlockTableRecord1 As BlockTableRecord
        Dim Block1 As BlockReference

        Dim Locatie_blocuri As String = Locatie.Locatie_blocuri
        Dim Locatie_blocuri_alternativ As String = Locatie.Locatie_blocuri_my_docs
        Dim Locatie1 As String = Locatie_fisier(Locatie_blocuri_alternativ & Nume_fisier, Locatie_blocuri & Nume_fisier)
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Locatie_blocuri & Nume_fisier) = False Then
            MsgBox("File " & Nume_fisier & " not found")
            Exit Function
        End If
        If Locatie1 = Locatie_blocuri & Nume_fisier Then
            If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Locatie.Locatie_blocuri_my_docs_) = False Then
                Microsoft.VisualBasic.FileIO.FileSystem.CreateDirectory(Locatie.Locatie_blocuri_my_docs_)
            End If
            If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Locatie_blocuri_alternativ & Nume_fisier) = False Then
                Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(Locatie_blocuri & Nume_fisier, Locatie_blocuri_alternativ & Nume_fisier)
            End If
            Locatie1 = Locatie_blocuri_alternativ & Nume_fisier
        End If

        If IO.File.Exists(Locatie1) = True Then
            Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                Using trans As Transaction = ThisDrawing.Database.TransactionManager.StartTransaction
                    BlockTable1 = trans.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                    Try
                        Using db As New Database(False, False)
                            'read block drawing do we need Lockdocument..?? Not Always..?
                            db.ReadDwgFile(Locatie1, System.IO.FileShare.Read, True, Nothing)
                            Using Trans2 As Transaction = ThisDrawing.TransactionManager.StartTransaction()
                                If BlockTable1.Has(NumeBlock) = False Then
                                    Dim idBTR As ObjectId = ThisDrawing.Database.Insert(NumeBlock, db, False)
                                    Trans2.Commit()
                                End If
                            End Using
                        End Using


                        'insert block 


                        If BlockTable1.Has(NumeBlock) Then
                            'block found, get instance for copying
                            BlockTableRecord1 = trans.GetObject(BlockTable1.Item(NumeBlock), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Else
                            MsgBox(vbLf & NumeBlock & " NOT FOUND")
                            Return Nothing
                            Exit Function
                        End If

                        trans.Commit()
                    Catch ex As System.Exception
                        Dim aex2 As New System.Exception("Error in inserting new block: " & NumeBlock & ": ", ex)
                        Throw aex2
                    Finally

                    End Try
                End Using
            End Using

            Return BlockTableRecord1
        Else
            MsgBox("Block file not found")
            Return Nothing
        End If



    End Function


    Public Function Get_Arrow_dimension_ID(ByVal NUME_variabila As String, ByVal NUME_ARROW As String) As ObjectId
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim OLD_VALUE As String = Application.GetSystemVariable(NUME_variabila)
            Application.SetSystemVariable(NUME_variabila, NUME_ARROW)
            If Not OLD_VALUE.Length = 0 Then Application.SetSystemVariable(NUME_variabila, OLD_VALUE)
            Dim ID1 As ObjectId
            Using Trans1 As Transaction = ThisDrawing.Database.TransactionManager.StartTransaction
                Dim BlockTable1 As BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                ID1 = BlockTable1(NUME_ARROW)
                Trans1.Commit()


            End Using

            Return ID1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function


    Public Function get_csf_point_from_chainage(ByRef Chainage As Double, ByRef Poly3d As Curve) As Point3d
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim Data_table1 As New System.Data.DataTable
        Data_table1.Columns.Add("TEXT325", GetType(DBText))
        Data_table1.Columns.Add("CHAINAGE_ON_POLY", GetType(Double))
        Data_table1.Columns.Add("CHAINAGE_PUBLISHED", GetType(Double))
        Dim Index1 As Double = 0
        'Dim Data_table2 As New System.Data.DataTable
        'Data_table2.Columns.Add("TEXT0", GetType(DBText))
        'Dim Index2 As Double = 0



        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

            For Each ObjID In BTrecord
                Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If TypeOf DBobject Is DBText Then
                    Dim Text1 As DBText = DBobject
                    If Text1.Layer = Poly3d.Layer Then
                        If Text1.Rotation > 3 * PI / 2 Then
                            Dim Chainage_text As String = Text1.TextString
                            If IsNumeric(Replace(Chainage_text, "+", "")) = True Then
                                Data_table1.Rows.Add()
                                Data_table1.Rows(Index1).Item("TEXT325") = Text1
                                Dim point_on_poly As Point3d = Poly3d.GetClosestPointTo(Text1.Position, Vector3d.ZAxis, False)
                                Data_table1.Rows(Index1).Item("CHAINAGE_ON_POLY") = Poly3d.GetDistAtPoint(point_on_poly)
                                Data_table1.Rows(Index1).Item("CHAINAGE_PUBLISHED") = CDbl(Replace(Chainage_text, "+", ""))

                                Index1 = Index1 + 1
                            End If

                        End If
                        If Text1.Rotation >= 0 And Text1.Rotation < PI / 4 Then
                            'Data_table2.Rows.Add()
                            'Data_table2.Rows(Index2).Item("TEXT0") = Text1
                            'Index2 = Index2 + 1
                        End If

                    End If
                End If
            Next

            Data_table1 = Sort_data_table(Data_table1, "CHAINAGE_ON_POLY")
            Dim Distanta1 As Double = 20
            Dim index_data As Double = 20.1


            If Data_table1.Rows.Count > 0 Then
                For i = 0 To Data_table1.Rows.Count - 1
                    If IsDBNull(Data_table1.Rows(i).Item("TEXT325")) = False Then
                        If IsDBNull(Data_table1.Rows(i).Item("CHAINAGE_PUBLISHED")) = False Then
                            If Abs(Data_table1.Rows(i).Item("CHAINAGE_PUBLISHED") - Chainage) < Distanta1 Then
                                Distanta1 = Abs(Data_table1.Rows(i).Item("CHAINAGE_PUBLISHED") - Chainage)
                                index_data = i
                            End If
                        End If
                    End If
                Next

                If Distanta1 < 20 And Not index_data = 20.1 Then
                    If IsDBNull(Data_table1.Rows(index_data).Item("CHAINAGE_ON_POLY")) = False And IsDBNull(Data_table1.Rows(index_data).Item("CHAINAGE_PUBLISHED")) = False Then
                        Dim diferenta As Double = Chainage - Data_table1.Rows(index_data).Item("CHAINAGE_PUBLISHED")


                        Return (Poly3d.GetPointAtDist(Data_table1.Rows(index_data).Item("CHAINAGE_ON_POLY") + diferenta))

                    End If
                Else
                    Return (Poly3d.GetPointAtDist(Chainage))
                End If

            Else
                Return (Poly3d.GetPointAtDist(Chainage))
            End If







        End Using



    End Function

    Public Function Stabileste_coloanele(ByVal String1 As String) As Integer
        String1 = String1.ToUpper
        String1 = Replace(String1, " ", "")
        If Len(String1) = 1 Then
            Return Asc(String1) - 64
            Exit Function
        End If

        If Len(String1) = 2 Then
            Dim St1, St2 As String
            St1 = Strings.Left(String1, 1)
            St2 = Strings.Right(String1, 1)
            Dim Val1, Val2 As Integer
            Val1 = Asc(St1) - 64
            Val2 = Asc(St2) - 64
            Return Val1 * 26 + Val2
        Else
            Return -1
        End If

    End Function
    Public Function RedefineBlock_from_browser_location(ByVal Nume_fisier_cu_path As String, ByVal NumeBlock As String) As BlockTableRecord
        Dim dlock As DocumentLock = Nothing
        Dim BlockTable1 As BlockTable
        Dim id As ObjectId
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
        Dim blockTableRec As BlockTableRecord
        Dim idBTR As ObjectId
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor

            'insert block and rename it
            Try
                Try
                    dlock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Catch ex As Exception
                    Dim aex As New System.Exception("Error locking document for InsertBlock: " & NumeBlock & ": ", ex)
                    Throw aex
                End Try
                BlockTable1 = trans.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                If BlockTable1.Has(NumeBlock) = True Then
                    Try
                        Dim Locatie1 As String = Nume_fisier_cu_path
                        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Nume_fisier_cu_path) = False Then
                            MsgBox("File " & Nume_fisier_cu_path & " not found")
                            Exit Function
                        End If
                        Dim ThisDrawing As Document = Application.DocumentManager.MdiActiveDocument
                        Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                            Using Database2 As New Database(False, False)
                                Database2.ReadDwgFile(Locatie1, System.IO.FileShare.Read, True, Nothing)
                                Using Trans1 As Transaction = ThisDrawing.TransactionManager.StartTransaction()
                                    Dim BlockTable2 As BlockTable = DirectCast(Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False), BlockTable)
                                    idBTR = ThisDrawing.Database.Insert(NumeBlock, Database2, True)
                                    Trans1.Commit()
                                End Using
                            End Using

                            If Not idBTR = ObjectId.Null Then
                                blockTableRec = trans.GetObject(idBTR, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False, True)

                                Dim ObjIdCol As ObjectIdCollection = blockTableRec.GetBlockReferenceIds(False, True)

                                For Each id1 As ObjectId In ObjIdCol
                                    Dim Bref As BlockReference = trans.GetObject(id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, False, True)
                                    Bref.RecordGraphicsModified(True)
                                Next
                            End If


                        End Using
                    Catch e As System.Exception

                        MsgBox(e.Message)
                    End Try

                End If ' ASTA E DE LA  If bt.Has THE BLOCK

                trans.Commit()

            Catch ex As System.Exception
                Dim aex2 As New System.Exception("Error in inserting new block: " & Nume_fisier_cu_path & ": ", ex)
                Throw aex2
            Finally
                If Not trans Is Nothing Then trans.Dispose()
                If Not dlock Is Nothing Then dlock.Dispose()
            End Try
        End Using
        Return blockTableRec
    End Function




    Public Function Intersect_on_both_operands(ByVal Curba1 As Curve, ByVal Curba2 As Curve) As Point3dCollection
        Dim Col_int As New Point3dCollection
        Dim Col_int_on_both_operands As New Point3dCollection
        Dim Col_int_on_both_operands_DUPLICATE As New Point3dCollection

        Curba1.IntersectWith(Curba2, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero)


        If Col_int.Count = 1 Then
            Return Col_int
        End If

        If Col_int.Count = 0 Then
            Return Col_int_on_both_operands
        End If

        If Col_int.Count > 1 Then
            If TypeOf Curba1 Is Polyline And TypeOf Curba2 Is Polyline Then
                For i = 0 To Col_int.Count - 1

                    Dim Pt1 As New Point3d
                    Pt1 = Col_int(i)

                    Try
                        Dim param_on_1 As Double
                        param_on_1 = Curba1.GetParameterAtPoint(Pt1)
                        Dim param_on_2 As Double
                        param_on_2 = Curba2.GetParameterAtPoint(Pt1)

                        If Col_int_on_both_operands_DUPLICATE.Contains(New Point3d(Round(Pt1.X, 4), Round(Pt1.Y, 4), Round(Pt1.Z, 4))) = False Then
                            Col_int_on_both_operands.Add(Pt1)
                            Col_int_on_both_operands_DUPLICATE.Add(New Point3d(Round(Pt1.X, 4), Round(Pt1.Y, 4), Round(Pt1.Z, 4)))
                        End If
                    Catch ex As Exception

                    End Try


                Next

                Return Col_int_on_both_operands

            Else
                Return Col_int
            End If




        End If





    End Function




    Public Function Angle_left_right(ByVal Poly2D As Polyline, ByVal Punct1 As Point3d, Optional output_just_LT_RT As Integer = 0) As String

        Dim Point_on_poly As Point3d = Poly2D.GetClosestPointTo(Punct1, Vector3d.ZAxis, True)
        Dim vector2 As Vector3d = Point_on_poly.GetVectorTo(Punct1)
        Dim Param1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly)
        Dim vector1 As Vector3d
        If Param1 > 0 Then
            If Param1 = Poly2D.NumberOfVertices - 1 Then
                vector1 = Poly2D.GetPoint3dAt(Param1 - 1).GetVectorTo(Poly2D.GetPoint3dAt(Param1))
            Else
                vector1 = Poly2D.GetPoint3dAt(Floor(Param1)).GetVectorTo(Poly2D.GetPoint3dAt(Ceiling(Param1)))
            End If

        Else
            vector1 = Poly2D.GetPoint3dAt(0).GetVectorTo(Poly2D.GetPoint3dAt(1))
        End If

        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

        Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
        Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

        Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI

        Dim LT_RT As String = ""


        If Bearing1 < 180 Then
            If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                LT_RT = " LT."
            Else
                LT_RT = " RT."
            End If
        Else
            If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                LT_RT = " RT."
            Else
                LT_RT = " LT."
            End If
        End If

        Dim AngleDMS As String = Floor(angle1) & "°"

        Dim Minute As String = Round((angle1 - Floor(angle1)) * 60, 0) & "'"

        If Round((angle1 - Floor(angle1)) * 60, 0) = 60 Then
            AngleDMS = Floor(angle1 + 1) & "°"
            Minute = "00'"
        End If

        If Len(Minute) = 2 Then Minute = "0" & Minute
        AngleDMS = AngleDMS & Minute & "00" & Chr(34)

        Dim String_DMS As String = AngleDMS & LT_RT
        If output_just_LT_RT > 0 Then
            Return LT_RT
        Else
            Return String_DMS
        End If




    End Function

    Public Sub Stretch_block(ByVal BR As BlockReference, ByVal Prop_name As String, ByVal Prop_value As Double)
        Dim pc As DynamicBlockReferencePropertyCollection = BR.DynamicBlockReferencePropertyCollection

        For Each prop As DynamicBlockReferenceProperty In pc
            If prop.PropertyName = Prop_name And prop.UnitsType = DynamicBlockReferencePropertyUnitsType.Distance Then
                prop.Value = Prop_value#
                Exit For
            End If
        Next


    End Sub

    Public Sub Add_to_clipboard_Data_table(ByVal Data_table As System.Data.DataTable)
        Dim sTR1 As String = ""




        If Data_table.Rows.Count > 0 Then

            For i = 0 To Data_table.Columns.Count - 1
                If i = 0 Then
                    sTR1 = Data_table.Columns(i).ColumnName
                Else
                    sTR1 = sTR1 & Chr(9) & Data_table.Columns(i).ColumnName
                End If
            Next
            For i = 0 To Data_table.Rows.Count - 1

                sTR1 = sTR1 & vbCrLf
                For j = 0 To Data_table.Columns.Count - 1
                    If IsDBNull(Data_table.Rows(i).Item(j)) = False Then


                        If j = 0 Then
                            sTR1 = sTR1 & Data_table.Rows(i).Item(j)
                        Else
                            sTR1 = sTR1 & Chr(9) & Data_table.Rows(i).Item(j)
                        End If
                    Else
                        If j = 0 Then
                            sTR1 = sTR1 & ""
                        Else
                            sTR1 = sTR1 & Chr(9) & ""
                        End If

                    End If
                Next

            Next

        End If


        My.Computer.Clipboard.SetText(sTR1)
    End Sub
    Public Sub UpdateBlock_with_multiple_atributes(ByVal NumeBlock As String, ByVal Database1 As Autodesk.AutoCAD.DatabaseServices.Database, ByVal Space1 As BlockTableRecord,
                                                        ByVal Colectie_nume_atribute As Specialized.StringCollection, Colectie_valori_atribute As Specialized.StringCollection)

        Dim BlockTable1 As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing

        Dim id As ObjectId

        Using Trans1 As Transaction = Database1.TransactionManager.StartTransaction

            Try

                BlockTable1 = Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                If BlockTable1.Has(NumeBlock) = True Then
                    Block_table_record1 = Trans1.GetObject(BlockTable1.Item(NumeBlock), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    For Each obid As ObjectId In Space1
                        Dim Block1 As BlockReference = TryCast(Trans1.GetObject(obid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockReference)
                        If IsNothing(Block1) = False Then

                            Dim Nume1 As String = ""

                            If Block1.AttributeCollection.Count > 0 Then
                                Dim BlockTrec As BlockTableRecord = Nothing
                                If Block1.IsDynamicBlock = True Then
                                    BlockTrec = Trans1.GetObject(Block1.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Nume1 = BlockTrec.Name
                                Else
                                    BlockTrec = Trans1.GetObject(Block1.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Nume1 = BlockTrec.Name
                                End If

                                If Nume1 = NumeBlock Then
                                    Block1.UpgradeOpen()


                                    Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                                    attColl = Block1.AttributeCollection
                                    For Each ID1 As ObjectId In attColl
                                        Dim ent As DBObject = Trans1.GetObject(ID1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                        If TypeOf ent Is AttributeReference Then

                                            Dim attref As AttributeReference = ent


                                            For i = 0 To Colectie_nume_atribute.Count - 1
                                                If attref.Tag = Colectie_nume_atribute(i) Then
                                                    Dim Valoare As String = Colectie_valori_atribute(i)
                                                    If Valoare = Nothing Then
                                                        Valoare = ""
                                                    End If

                                                    If attref.IsMTextAttribute = False Then
                                                        attref.TextString = Valoare
                                                    Else
                                                        attref.MTextAttribute.Contents = Valoare
                                                    End If

                                                End If
                                            Next



                                        End If
                                    Next



                                End If
                            End If
                        End If
                    Next
                    Trans1.Commit()

                End If ' ASTA E DE LA  If bt.Has THE BLOCK
            Catch ex As System.Exception
                MsgBox(ex.Message)
            Finally
                If Not Trans1 Is Nothing Then Trans1.Dispose()

            End Try
        End Using

    End Sub


    Public Function Transfer_datatable_to_new_excel_spreadsheet(dt1 As System.Data.DataTable)
        If IsNothing(dt1) = False Then
            If dt1.Rows.Count > 0 Then
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = get_new_worksheet_from_Excel()
                W1.Cells.NumberFormat = "@"

                Dim maxRows As Integer = dt1.Rows.Count
                Dim maxCols As Integer = dt1.Columns.Count

                Dim range As Microsoft.Office.Interop.Excel.Range = W1.Range(W1.Cells(2, 1), W1.Cells(maxRows + 1, maxCols))

                Dim values(maxRows, maxCols) As Object

                For row = 0 To maxRows - 1
                    For col = 0 To maxCols - 1
                        If IsDBNull(dt1.Rows(row).Item(col)) = False Then
                            values(row, col) = dt1.Rows(row).Item(col)
                        End If
                    Next
                Next

                range.Value2 = values

                For i = 0 To dt1.Columns.Count - 1
                    W1.Cells(1, i + 1).value2 = dt1.Columns(i).ColumnName
                Next
            End If
        End If




    End Function


    Public Function zoom_to_Point(pt As Point3d)


        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Try
            Using Lock1 As Autodesk.AutoCAD.ApplicationServices.DocumentLock = ThisDrawing.LockDocument()
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction()

                    Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim factor1 As Double = 200
                    Dim minx As New Point3d(pt.X - factor1, pt.Y - factor1, 0)
                    Dim maxx As New Point3d(pt.X + factor1, pt.Y + factor1, 0)
                    Using GraphicsManager As Autodesk.AutoCAD.GraphicsSystem.Manager = ThisDrawing.GraphicsManager
                        Dim Cvport As Integer = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"))
                        Dim kd As Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor = New Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor()
                        kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"))
                        Dim view As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.ObtainAcGsView(Cvport, kd)
                        If IsNothing(view) = False Then
                            Using view
                                view.ZoomExtents(minx, maxx)
                                view.Zoom(0.95)
                                GraphicsManager.SetViewportFromView(Cvport, view, True, True, False)
                            End Using
                        End If


                        Trans1.Commit()

                    End Using
                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        Return Nothing
    End Function

End Module
