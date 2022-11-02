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
Imports ACSMCOMPONENTS20Lib
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
                Or diskPropertyB.Value.ToString() = "18399D24" _
                Or diskPropertyB.Value.ToString() = "8C040338" Then
                IsOK = True
                Return IsOK
                Exit Function
            End If

            Try
                If Environment.GetEnvironmentVariable("USERDNSDOMAIN").ToString.ToUpper = "HMMG.CC" Or Environment.GetEnvironmentVariable("USERDNSDOMAIN").ToString.ToLower() = "mottmac.group.int" Then
                    IsOK = True
                Else
                    If Today.Year = 2017 And Today.Month <= 12 And Today.Day < 25 Then
                        If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists("\\inandpz01\projects") = True Then
                            IsOK = True
                            Return IsOK
                            Exit Function
                        End If

                        'IsOK = True
                    End If
                End If

            Catch ex As System.Exception
                If Today.Year = 2017 And Today.Month <= 12 And Today.Day < 25 Then
                    If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists("\\inandpz01\projects") = True Then
                        IsOK = True
                        Return IsOK
                        Exit Function
                    End If
                End If
            End Try

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
    Public Function Get_active_worksheet_from_Excel() As Microsoft.Office.Interop.Excel.Worksheet
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook
        Try
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch ex As Exception
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

    Public Function Get_NEW_worksheet_from_Excel() As Microsoft.Office.Interop.Excel.Worksheet
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook
        Try
            Excel1 = New Microsoft.Office.Interop.Excel.Application
            Excel1.Visible = True
            Excel1.Workbooks.Add()
            Workbook1 = Excel1.ActiveWorkbook
        Catch ERROARE As Exception
            MsgBox(ERROARE.Message)
        End Try
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Return Workbook1.ActiveSheet
    End Function

    Public Function GET_Bearing_rad(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double


        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

        Return New Point3d(x1, y1, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)



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

    Public Function InsertBlock_with_multiple_atributes(ByVal Nume_fisier As String, ByVal NumeBlock As String, _
                                                            ByVal Insertion_point As Point3d, ByVal Scale_xyz As Double, ByVal Spatiu As BlockTableRecord, _
                                                            ByVal Layer1 As String, _
                                                            ByVal Colectie_nume_atribute As Specialized.StringCollection, Colectie_valori_atribute As Specialized.StringCollection) As BlockReference
        Dim dlock As DocumentLock = Nothing
        Dim BlockTable1 As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing
        Dim br As BlockReference
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
                        Dim Locatie1 As String
                        Dim Locatie_blocuri As String = Locatie.Locatie_blocuri
                        Dim Locatie_blocuri_alternativ As String = Locatie.Locatie_blocuri_my_docs
                        Dim Locatie_de_cautat As String = Locatie_blocuri & Nume_fisier

                        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Nume_fisier) = True Then
                            If Nume_fisier.ToUpper.Contains("C:\") = True Then
                                Locatie1 = Nume_fisier
                                GoTo l123
                            End If
                        End If

                        Locatie1 = Locatie_fisier(Locatie_blocuri_alternativ & Nume_fisier, Locatie_blocuri & Nume_fisier)
                        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Locatie_de_cautat) = False Then
                            If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Locatie_blocuri_alternativ & Nume_fisier) = False Then
                                MsgBox("File " & Nume_fisier & " not found")
                                Exit Function
                            End If
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

l123:


                        Dim ThisDrawing As Document = Application.DocumentManager.MdiActiveDocument

                        Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                            Using Database2 As New Database(False, False)
                                'read block drawing do we need Lockdocument..?? Not Always..?
                                Database2.ReadDwgFile(Locatie1, System.IO.FileShare.Read, True, Nothing)
                                Using Trans1 As Transaction = ThisDrawing.TransactionManager.StartTransaction()

                                    Dim BlockTable2 As BlockTable = DirectCast(Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False), BlockTable)

                                    Dim idBTR As ObjectId = ThisDrawing.Database.Insert(NumeBlock, Database2, False)
                                    Trans1.Commit()

                                End Using
                            End Using
                        End Using
                    Catch e As System.Exception

                        MsgBox(e.Message)
                    End Try
                    GoTo 12345
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

    Public Function x1_for_arc_leader(ByVal x0 As Double, ByVal y0 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal Inaltimea As Double) As Double

        Dim Length1 As Double = (Inaltimea) / 500
        Dim Dist1 As Double = GET_distanta_Double(x0, y0, x2, y2)


        Dim x1, y1 As Double
        Dim Dx, Dy As Double

        Dim Bear1 As Double = GET_Bearing_rad(x0, y0, x2, y2)

        If Bear1 < PI / 2 Then
            Dx = Length1 * Cos(Bear1)
            Dy = Length1 * Sin(Bear1)
            x1 = x0 + Dx
            y1 = y0 + Dy
        End If

        If Bear1 = PI / 2 Then
            Dx = 0
            Dy = Length1
            x1 = x0
            y1 = y0 + Dy
        End If

        If Bear1 > PI / 2 And Bear1 < PI Then
            Dx = Length1 * Cos(PI - Bear1)
            Dy = Length1 * Sin(PI - Bear1)
            x1 = x0 - Dx
            y1 = y0 + Dy
        End If

        If Bear1 = PI Then
            Dx = Length1
            Dy = 0
            x1 = x0 - Dx
            y1 = y0
        End If

        If Bear1 > PI And Bear1 < 3 * PI / 2 Then
            Dx = Length1 * Cos(Bear1 - PI)
            Dy = Length1 * Sin(Bear1 - PI)
            x1 = x0 - Dx
            y1 = y0 - Dy
        End If

        If Bear1 = 3 * PI / 2 Then
            Dx = 0
            Dy = Length1
            x1 = x0
            y1 = y0 - Dy
        End If

        If Bear1 > 3 * PI / 2 And Bear1 < 2 * PI Then
            Dx = Length1 * Cos(2 * PI - Bear1)
            Dy = Length1 * Sin(2 * PI - Bear1)
            x1 = x0 + Dx
            y1 = y0 - Dy
        End If

        If Bear1 = 2 * PI Or Bear1 = 0 Then
            Dx = Length1
            Dy = 0
            x1 = x0 + Dx
            y1 = y0
        End If

        Return x1

    End Function

    Public Function y1_for_arc_leader(ByVal x0 As Double, ByVal y0 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal Inaltimea As Double) As Double

        Dim Length1 As Double = (Inaltimea) / 500
        Dim Dist1 As Double = GET_distanta_Double(x0, y0, x2, y2)


        Dim x1, y1 As Double
        Dim Dx, Dy As Double

        Dim Bear1 As Double = GET_Bearing_rad(x0, y0, x2, y2)

        If Bear1 < PI / 2 Then
            Dx = Length1 * Cos(Bear1)
            Dy = Length1 * Sin(Bear1)
            x1 = x0 + Dx
            y1 = y0 + Dy
        End If

        If Bear1 = PI / 2 Then
            Dx = 0
            Dy = Length1
            x1 = x0
            y1 = y0 + Dy
        End If

        If Bear1 > PI / 2 And Bear1 < PI Then
            Dx = Length1 * Cos(PI - Bear1)
            Dy = Length1 * Sin(PI - Bear1)
            x1 = x0 - Dx
            y1 = y0 + Dy
        End If

        If Bear1 = PI Then
            Dx = Length1
            Dy = 0
            x1 = x0 - Dx
            y1 = y0
        End If

        If Bear1 > PI And Bear1 < 3 * PI / 2 Then
            Dx = Length1 * Cos(Bear1 - PI)
            Dy = Length1 * Sin(Bear1 - PI)
            x1 = x0 - Dx
            y1 = y0 - Dy
        End If

        If Bear1 = 3 * PI / 2 Then
            Dx = 0
            Dy = Length1
            x1 = x0
            y1 = y0 - Dy
        End If

        If Bear1 > 3 * PI / 2 And Bear1 < 2 * PI Then
            Dx = Length1 * Cos(2 * PI - Bear1)
            Dy = Length1 * Sin(2 * PI - Bear1)
            x1 = x0 + Dx
            y1 = y0 - Dy
        End If

        If Bear1 = 2 * PI Or Bear1 = 0 Then
            Dx = Length1
            Dy = 0
            x1 = x0 + Dx
            y1 = y0
        End If

        Return y1

    End Function

    Public Function Bulge_for_arc_leader(ByVal x0 As Double, ByVal y0 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal Inaltimea As Double) As Double

        Dim x1, y1 As Double
        x1 = x1_for_arc_leader(x0, y0, x2, y2, Inaltimea)
        y1 = y1_for_arc_leader(x0, y0, x2, y2, Inaltimea)

        Dim B1 As Double = GET_Bearing_rad(x1, y1, x0, y0)

        Dim B2 As Double = GET_Bearing_rad(x1, y1, x3, y3)

        Dim Angle1 As Double

        If B1 <= PI / 2 Then
            If B2 >= B1 And B2 <= B1 + PI Then
                Angle1 = B2 - B1
                ' - bulge
            Else
                Angle1 = 2 * PI - (B2 - B1)
                ' + bulge
            End If
        End If

        If B1 > PI / 2 And B1 <= PI Then
            If B2 >= B1 And B2 <= B1 + PI Then
                Angle1 = B2 - B1
                ' - bulge
            Else
                Angle1 = 2 * PI - (B2 - B1)
                ' + bulge
            End If
        End If

        If B1 > PI And B1 <= 3 * PI / 2 Then
            If B2 <= B1 And B2 >= B1 - PI Then
                Angle1 = -(B2 - B1)
                ' + bulge
            Else
                Angle1 = 2 * PI + (B2 - B1)
                ' - bulge
            End If
        End If

        If B1 > 3 * PI / 2 And B1 <= 2 * PI Then
            If B2 <= B1 And B2 >= B1 - PI Then
                Angle1 = -(B2 - B1)
                ' + bulge
            Else
                Angle1 = 2 * PI + (B2 - B1)
                ' - bulge
            End If
        End If

        If Angle1 >= 2 * PI Then Angle1 = Angle1 - 2 * PI


        Dim a2 As Double
        Dim Unghi_la_centru As Double


        If Angle1 >= PI / 2 Then
            a2 = Angle1 - PI / 2
            Unghi_la_centru = PI - 2 * a2
        Else
            a2 = PI / 2 - Angle1
            Unghi_la_centru = 2 * PI - (PI - 2 * a2)
        End If

        Dim Bulge1 As Double
        Bulge1 = Tan(Unghi_la_centru / 4)

        If B1 <= PI / 2 Then
            If B2 >= B1 And B2 <= B1 + PI Then
                Bulge1 = -Bulge1
            End If
        End If

        If B1 > PI / 2 And B1 <= PI Then
            If B2 >= B1 And B2 <= B1 + PI Then
                Bulge1 = -Bulge1
            End If
        End If

        If B1 > PI And B1 <= 3 * PI / 2 Then
            If B2 <= B1 And B2 >= B1 - PI Then
            Else
                Bulge1 = -Bulge1
            End If
        End If

        If B1 > 3 * PI / 2 And B1 <= 2 * PI Then
            If B2 <= B1 And B2 >= B1 - PI Then
            Else
                Bulge1 = -Bulge1
            End If
        End If

        Return Bulge1

    End Function

    Public Function GET_distanta_Double(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
        Return ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
    End Function

    Public Function Quadrant_bearings(ByVal Radians As Double) As String
        Dim Angle1 As Double = Radians * 180 / PI

        Dim Prefix1 As String = ""
        Dim Suffix1 As String = ""

        If Angle1 >= 360 Then
            Angle1 = Angle1 - 360
        End If

        If Angle1 <= 90 Then
            Prefix1 = "N"
            Suffix1 = "E"
            Angle1 = 90 - Angle1
        End If

        If Angle1 > 90 And Angle1 <= 180 Then
            Prefix1 = "N"
            Suffix1 = "W"
            Angle1 = Angle1 - 90
        End If

        If Angle1 > 180 And Angle1 <= 270 Then
            Prefix1 = "S"
            Suffix1 = "W"
            Angle1 = 270 - Angle1
        End If

        If Angle1 > 270 Then
            Prefix1 = "S"
            Suffix1 = "E"
            Angle1 = Angle1 - 270
        End If



        Dim Degree As Integer = Floor(Angle1)
        Dim Min1 As Integer = Round((Angle1 - Floor(Angle1)) * 60, 0)
        ' Dim Min1 As Integer = Floor((Angle1 - Floor(Angle1)) * 60)
        'Dim Sec1 As Integer = Round(((Angle1 - Degree) * 60 - Min1) * 60, 0)
        'If Sec1 = 60 Then
        'Sec1 = 0
        'Min1 = Min1 + 1
        'End If
        If Min1 = 60 Then
            Degree = Degree + 1
            Min1 = 0
        End If

        Dim Minute As String = Min1.ToString
        'If Len(Minute) = 1 Then Minute = "0" & Minute
        'Dim Second As String = Sec1.ToString
        'If Len(Second) = 1 Then Second = "0" & Second

        Return Prefix1 & " " & Degree.ToString & "°" & Minute & "' " & Suffix1

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
                    Combo_atributes.Items.Add("")


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

    Public Function Bearing_dist_calc_with_min_deflection(ByVal min_angle_rad As Double, ByVal Poly1 As Polyline, ByVal Param_picked As Double) As Polyline
        Dim Start3D As New Point3d
        Dim End3D As New Point3d

        Dim Param1, Param2 As Double
        If Param_picked = Round(Param_picked, 0) Then
            If Param_picked = 0 Then
                Param1 = 0
                Param2 = 1
            ElseIf Param_picked = Poly1.NumberOfVertices - 1 Then
                Param1 = Poly1.NumberOfVertices - 2
                Param2 = Poly1.NumberOfVertices - 1
            Else
                Param1 = Param_picked
                Param2 = Param_picked + 1
            End If
        Else
            Param1 = Floor(Param_picked)
            Param2 = Ceiling(Param_picked)
        End If

        Dim kk As Integer = 0
        Dim IS_Smaller1 As Boolean = True
        Do While IS_Smaller1 = True
            If Param1 - kk >= 0 Then
                Dim Pt_0 As New Point3d
                Pt_0 = Poly1.GetPointAtParameter(Param1 - kk)

                Dim Pt_1 As New Point3d
                Pt_1 = Poly1.GetPointAtParameter(Param1 - kk + 1)

                Dim Vector1 As New Vector3d
                Vector1 = New Point3d(Pt_0.X, Pt_0.Y, 0).GetVectorTo(New Point3d(Pt_1.X, Pt_1.Y, 0))

                Dim Pt_2 As New Point3d

                If Param1 - kk - 1 >= 0 Then
                    Pt_2 = Poly1.GetPointAtParameter(Param1 - kk - 1)
                    Dim Vector0 As New Vector3d
                    Vector0 = New Point3d(Pt_2.X, Pt_2.Y, 0).GetVectorTo(New Point3d(Pt_0.X, Pt_0.Y, 0))

                    Dim angle1 As Double = Vector0.GetAngleTo(Vector1)

                    If angle1 < min_angle_rad Then
                        kk = kk + 1
                    Else
                        Start3D = Poly1.GetPointAtParameter(Param1 - kk)
                        IS_Smaller1 = False
                    End If
                Else
                    Start3D = Poly1.StartPoint
                    IS_Smaller1 = False
                End If
            Else
                Start3D = Poly1.StartPoint
                IS_Smaller1 = False
            End If
        Loop

        Dim jj As Integer = 0
        Dim IS_Smaller2 As Boolean = True

        Do While IS_Smaller2 = True
            If Param2 + jj <= Poly1.NumberOfVertices - 1 Then
                Dim Pt_0 As New Point3d
                Pt_0 = Poly1.GetPointAtParameter(Param2 + jj)

                Dim Pt_1 As New Point3d
                Pt_1 = Poly1.GetPointAtParameter(Param2 + jj - 1)

                Dim Vector0 As New Vector3d
                Vector0 = New Point3d(Pt_1.X, Pt_1.Y, 0).GetVectorTo(New Point3d(Pt_0.X, Pt_0.Y, 0))

                Dim Pt_2 As New Point3d

                If Param2 + jj + 1 <= Poly1.NumberOfVertices - 1 Then
                    Pt_2 = Poly1.GetPointAtParameter(Param2 + jj + 1)
                    Dim Vector1 As New Vector3d
                    Vector1 = New Point3d(Pt_0.X, Pt_0.Y, 0).GetVectorTo(New Point3d(Pt_2.X, Pt_2.Y, 0))

                    Dim angle1 As Double = Vector0.GetAngleTo(Vector1)

                    If angle1 < min_angle_rad Then
                        jj = jj + 1
                    Else
                        End3D = Poly1.GetPointAtParameter(Param2 + jj)
                        IS_Smaller2 = False
                    End If
                Else
                    End3D = Poly1.EndPoint
                    IS_Smaller2 = False
                End If
            Else
                End3D = Poly1.EndPoint
                IS_Smaller2 = False
            End If
        Loop
        Dim Poly2 As New Polyline
        Dim Par_1 As Integer = Poly1.GetParameterAtPoint(Start3D)
        Dim Par_2 As Integer = Poly1.GetParameterAtPoint(End3D)
        Dim IDX As Integer = 0
        For i = Par_1 To Par_2
            Poly2.AddVertexAt(IDX, Poly1.GetPoint2dAt(i), Poly1.GetBulgeAt(i), Poly1.GetStartWidthAt(i), Poly1.GetEndWidthAt(i))
            IDX = IDX + 1
        Next
        Return Poly2

    End Function

    Public Function Directie_offset(ByVal Curba1 As Curve, ByVal Point1 As Point3d) As Integer
        Using trans As Transaction = Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction

            Dim Point_on_curve As Point3d = Curba1.GetClosestPointTo(Point1, Vector3d.ZAxis, False)
            Dim pln = New Plane(New Point3d(0, 0, 0), New Vector3d(0, 0, 1))
            Dim ptAlongObj As Point3d
            Try
                ptAlongObj = Curba1.GetPointAtDist(Curba1.GetDistAtPoint(Point_on_curve) + 1)
            Catch ex As Exception
                ptAlongObj = Curba1.EndPoint
            End Try
            Dim vecOnObj As Vector3d = ptAlongObj - Point_on_curve
            Dim vecToPoint As Vector3d = Point1 - Point_on_curve

            Dim angAlongObj As Double = vecOnObj.AngleOnPlane(pln)

            Dim ang As Double = vecToPoint.AngleOnPlane(pln)
            If ang < angAlongObj Then ang += Math.PI * 2


            If angAlongObj + Math.PI < ang Then
                If TypeOf Curba1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Return -1
                Else
                    Return 1
                End If
                Return 1
            Else
                If TypeOf Curba1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Return 1
                Else
                    Return -1
                End If
            End If


        End Using

    End Function

    Public Function LockDatabase(ByRef database As AcSmDatabase, _
                          ByVal lockFlag As Boolean) As Boolean
        Dim dbLock As Boolean = False

        ' If lockFalg equals True then attempt to lock the database, otherwise
        ' attempt to unlock it.
        If lockFlag = True And _
           database.GetLockStatus() = AcSmLockStatus.AcSmLockStatus_UnLocked Then
            database.LockDb(database)
            dbLock = True
        ElseIf lockFlag = False And _
            database.GetLockStatus = AcSmLockStatus.AcSmLockStatus_Locked_Local Then
            database.UnlockDb(database)
            dbLock = True
        Else
            dbLock = False
        End If

        LockDatabase = dbLock
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
                            sTR1 = sTR1 & Data_table.Rows(i).Item(j).ToString
                        Else
                            sTR1 = sTR1 & Chr(9) & Data_table.Rows(i).Item(j).ToString
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
    Public Function Transfer_datatable_to_new_excel_spreadsheet(dt1 As System.Data.DataTable)
        If IsNothing(dt1) = False Then
            If dt1.Rows.Count > 0 Then
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_NEW_worksheet_from_Excel()
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
End Module
