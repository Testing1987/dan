
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
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
    Public Function InsertBlock_with_multiple_atributes(ByVal Nume_fisier As String, ByVal NumeBlock As String, _
                                                          ByVal Insertion_point As Point3d, ByVal Scale_xyz As Double, ByVal Spatiu As BlockTableRecord, _
                                                          ByVal Layer1 As String, _
                                                          ByVal Colectie_nume_atribute As Specialized.StringCollection, Colectie_valori_atribute As Specialized.StringCollection) As BlockReference
        Dim Lock1 As DocumentLock = Nothing
        Dim BlockTable1 As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing
        Dim Block1 As BlockReference
        Dim id As ObjectId
        Dim Database1 As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = Database1.TransactionManager.StartTransaction
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor

            'insert block and rename it
            Try
                Try
                    Lock1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Catch ex As Exception
                    Dim aex As New System.Exception("Error locking document for InsertBlock: " & NumeBlock & ": ", ex)
                    Throw aex
                End Try
                BlockTable1 = trans.GetObject(Database1.BlockTableId, autodesk.autocad.databaseservices.openmode.ForWrite)
12345:
                If BlockTable1.Has(NumeBlock) = True Then
                    Block_table_record1 = trans.GetObject(BlockTable1.Item(NumeBlock), autodesk.autocad.databaseservices.openmode.ForRead)


                Else
                    Try
                        Dim Locatie1 As String
                      

                        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Nume_fisier) = True Then
                            Locatie1 = Nume_fisier
                        Else
                            MsgBox("File " & Nume_fisier & " not found")
                            Exit Function
                        End If





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

                Spatiu = trans.GetObject(Spatiu.ObjectId, autodesk.autocad.databaseservices.openmode.ForWrite)
                'Set the Attribute Value
                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                Dim ent As Entity
                Dim Block_table_record1enum As BlockTableRecordEnumerator
                Block1 = New BlockReference(Insertion_point, Block_table_record1.ObjectId)
                Block1.Layer = Layer1
                Block1.ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(Scale_xyz, Scale_xyz, Scale_xyz)

                Spatiu.AppendEntity(Block1)
                trans.AddNewlyCreatedDBObject(Block1, True)
                attColl = Block1.AttributeCollection
                Block_table_record1enum = Block_table_record1.GetEnumerator
                While Block_table_record1enum.MoveNext
                    ent = Block_table_record1enum.Current.GetObject(autodesk.autocad.databaseservices.openmode.ForWrite)
                    If TypeOf ent Is AttributeDefinition Then
                        Dim attdef As AttributeDefinition = ent
                        Dim attref As New AttributeReference
                        attref.SetAttributeFromBlock(attdef, Block1.BlockTransform)

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
                If Not Lock1 Is Nothing Then Lock1.Dispose()
            End Try
        End Using
        Return Block1
    End Function

    Public Function InsertBlock_with_multiple_atributes_background(ByVal Nume_fisier As String, ByVal NumeBlock As String, ByVal Database1 As Autodesk.AutoCAD.DatabaseServices.Database, _
                                                          ByVal Insertion_point As Point3d, ByVal Scale_xyz As Double, ByVal Spatiu As BlockTableRecord, _
                                                          ByVal Layer1 As String, _
                                                          ByVal Colectie_nume_atribute As Specialized.StringCollection, Colectie_valori_atribute As Specialized.StringCollection) As BlockReference

        Dim BlockTable1 As BlockTable
        Dim Block_table_record1 As BlockTableRecord = Nothing
        Dim Block1 As BlockReference
        Dim id As ObjectId

        Using Trans1 As Transaction = Database1.TransactionManager.StartTransaction


            'insert block and rename it
            Try

                BlockTable1 = Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
12345:
                If BlockTable1.Has(NumeBlock) = True Then
                    Block_table_record1 = Trans1.GetObject(BlockTable1.Item(NumeBlock), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                Else
                    Try
                        

                        If IO.File.Exists(Nume_fisier) = False Then
                            MsgBox("File " & Nume_fisier & " not found")
                            Return Nothing
                        End If






                        Using Database2 As New Database(False, False)
                            'read block drawing do we need Lockdocument..?? Not Always..?
                            Database2.ReadDwgFile(Nume_fisier, System.IO.FileShare.Read, True, Nothing)
                            Using Trans2 As Transaction = Database1.TransactionManager.StartTransaction()

                                Dim BlockTable2 As BlockTable = DirectCast(Trans2.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False), BlockTable)

                                Dim idBTR As ObjectId = Database1.Insert(NumeBlock, Database2, False)
                                Trans2.Commit()

                            End Using
                        End Using

                    Catch e As System.Exception

                        MsgBox(e.Message)
                    End Try
                    GoTo 12345
                End If ' ASTA E DE LA  If bt.Has THE BLOCK

                Spatiu = Trans1.GetObject(Spatiu.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                'Set the Attribute Value
                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                Dim ent As Entity
                Dim Block_table_record1enum As BlockTableRecordEnumerator
                Block1 = New BlockReference(Insertion_point, Block_table_record1.ObjectId)
                Block1.Layer = Layer1
                Block1.ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(Scale_xyz, Scale_xyz, Scale_xyz)

                Spatiu.AppendEntity(Block1)
                Trans1.AddNewlyCreatedDBObject(Block1, True)
                attColl = Block1.AttributeCollection
                Block_table_record1enum = Block_table_record1.GetEnumerator
                While Block_table_record1enum.MoveNext
                    ent = Block_table_record1enum.Current.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    If TypeOf ent Is AttributeDefinition Then
                        Dim attdef As AttributeDefinition = ent
                        Dim attref As New AttributeReference
                        attref.SetAttributeFromBlock(attdef, Block1.BlockTransform)

                        For i = 0 To Colectie_nume_atribute.Count - 1
                            If attref.Tag = Colectie_nume_atribute(i) Then
                                Dim Valoare As String = ""
                                If IsNothing(Colectie_valori_atribute(i)) = False Then
                                    Valoare = Colectie_valori_atribute(i)
                                End If


                                If attref.IsMTextAttribute = False Then
                                    attref.TextString = Valoare
                                Else
                                    attref.MTextAttribute.Contents = Valoare
                                End If

                            End If
                        Next


                        If IsNothing(attref) = False Then
                            attColl.AppendAttribute(attref)
                            Trans1.AddNewlyCreatedDBObject(attref, True)
                        End If



                    End If
                End While
                Trans1.Commit()

            Catch ex As System.Exception
                Dim aex2 As New System.Exception("Error in inserting new block: " & Nume_fisier & ": ", ex)
                Throw aex2
            Finally
                If Not Trans1 Is Nothing Then Trans1.Dispose()

            End Try
        End Using
        Return Block1
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
    Public Function GET_Bearing_rad(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d
        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)
        Return New Point3d(x1, y1, 0).GetVectorTo(New Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent)
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
    Public Sub Stretch_block(ByVal BR As BlockReference, ByVal Prop_name As String, ByVal Prop_value As Double)
        Using pc As DynamicBlockReferencePropertyCollection = BR.DynamicBlockReferencePropertyCollection
            For Each prop As DynamicBlockReferenceProperty In pc
                If prop.PropertyName = Prop_name And prop.UnitsType = DynamicBlockReferencePropertyUnitsType.Distance Then
                    prop.Value = Prop_value#
                    Exit For
                End If
            Next
        End Using
    End Sub
    Public Function Get_distance1_block(ByVal BR As BlockReference) As Double
        Using pc As DynamicBlockReferencePropertyCollection = BR.DynamicBlockReferencePropertyCollection
            For Each prop As DynamicBlockReferenceProperty In pc
                If prop.PropertyName = "Distance1" And prop.UnitsType = DynamicBlockReferencePropertyUnitsType.Distance Then
                    Return prop.Value
                    Exit For
                End If
            Next
        End Using
    End Function

    Public Sub change_visibility_block(ByVal BR As BlockReference, ByVal Visibility_name As String)
        Using pc As DynamicBlockReferencePropertyCollection = BR.DynamicBlockReferencePropertyCollection
            For Each prop As DynamicBlockReferenceProperty In pc
                If prop.PropertyName = "Visibility1" And prop.PropertyTypeCode = 5 Then
                    prop.Value = Visibility_name
                    Exit For
                End If
            Next
        End Using
    End Sub

    Public Function Get_active_worksheet_from_Excel() As Microsoft.Office.Interop.Excel.Worksheet
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook
        Try
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch ex As System.Exception
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

    Public Function Get_active_worksheet_from_Excel_with_error() As Microsoft.Office.Interop.Excel.Worksheet
        Dim Excel1 As Microsoft.Office.Interop.Excel.Application
        Dim Workbook1 As Microsoft.Office.Interop.Excel.Workbook
        Try
            Excel1 = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch ex As System.SystemException
            Return Nothing
        Finally
            'Excel1.ActiveWindow.DisplayGridlines = True
            If Excel1.Workbooks.Count = 0 Then Excel1.Workbooks.Add()
            If Excel1.Visible = False Then Excel1.Visible = True
            Workbook1 = Excel1.ActiveWorkbook
        End Try
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
        Catch ERROARE As System.SystemException
            MsgBox(ERROARE.Message)
            Return Nothing
            Exit Function
        End Try
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Return Workbook1.ActiveSheet
    End Function
    Public Function Incarca_existing_Atributes_to_combobox(ByVal BlockName As String, ByVal Combo_atributes As Windows.Forms.ComboBox)
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using LOCK1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Block_table As Autodesk.AutoCAD.DatabaseServices.BlockTable = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Combo_atributes.Items.Clear()
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
                            If Block1.Name.Contains("*") = False Then
                                Combo_Blocks_with_atributes.Items.Add(Block1.Name)
                            End If
                        End If
                    Next
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
                        If Block1.Name.Contains("*") = False Then
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

    Public Function Sort_data_table(ByVal Datatable1 As System.Data.DataTable, ByVal Column1 As String) As System.Data.DataTable
        If Datatable1.Columns.Contains(Column1) = True Then
            Datatable1.DefaultView.Sort = Column1 & " ASC"
            Dim Sorted_DT As New System.Data.DataTable
            Sort_data_table = Datatable1.DefaultView.ToTable
            Return Sort_data_table
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SynchronizeAttributes_db_diferit(target As BlockTableRecord, Tr As Transaction)
        If target Is Nothing Then
            Throw New ArgumentNullException("btr")
        End If
        If Tr Is Nothing Then
            Throw New Exception(ErrorStatus.NoActiveTransactions)
        End If
        Dim attDefClass As RXClass = RXClass.GetClass(GetType(AttributeDefinition))
        Dim attDefs As New List(Of AttributeDefinition)()
        For Each id As ObjectId In target
            If id.ObjectClass = attDefClass Then
                Dim attDef As AttributeDefinition = DirectCast(Tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeDefinition)
                attDefs.Add(attDef)
            End If
        Next
        For Each id As ObjectId In target.GetBlockReferenceIds(True, False)
            Dim br As BlockReference = DirectCast(Tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), BlockReference)
            br.ResetAttributes(attDefs)
        Next
        If target.IsDynamicBlock Then
            For Each id As ObjectId In target.GetAnonymousBlockIds()
                Dim btr1 As BlockTableRecord = DirectCast(Tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTableRecord)
                For Each brId As ObjectId In btr1.GetBlockReferenceIds(True, False)
                    Dim br As BlockReference = DirectCast(Tr.GetObject(brId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), BlockReference)
                    br.ResetAttributes_db_diferit(attDefs)
                Next
            Next
        End If
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Private Sub ResetAttributes(br As BlockReference, attDefs As List(Of AttributeDefinition))
        Dim tm As Autodesk.AutoCAD.DatabaseServices.TransactionManager = br.Database.TransactionManager
        Dim attValues As New Dictionary(Of String, String)()
        For Each id As ObjectId In br.AttributeCollection
            If Not id.IsErased Then
                Dim attRef As AttributeReference = DirectCast(tm.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite), AttributeReference)
                attValues.Add(attRef.Tag, If(attRef.IsMTextAttribute, attRef.MTextAttribute.Contents, attRef.TextString))
                attRef.Erase()
            End If
        Next
        For Each attDef As AttributeDefinition In attDefs
            Dim attRef As New AttributeReference()
            attRef.SetAttributeFromBlock(attDef, br.BlockTransform)
            If attValues IsNot Nothing AndAlso attValues.ContainsKey(attDef.Tag) Then
                attRef.TextString = attValues(attDef.Tag.ToUpper())
            End If
            br.AttributeCollection.AppendAttribute(attRef)
            tm.AddNewlyCreatedDBObject(attRef, True)
        Next
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Private Sub ResetAttributes_db_diferit(br As BlockReference, attDefs As List(Of AttributeDefinition))
        Dim tm As Autodesk.AutoCAD.DatabaseServices.TransactionManager = br.Database.TransactionManager
        Dim attValues As New Dictionary(Of String, String)()
        For Each id As ObjectId In br.AttributeCollection
            If Not id.IsErased Then
                Dim attRef As AttributeReference = DirectCast(tm.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite), AttributeReference)
                attValues.Add(attRef.Tag, If(attRef.IsMTextAttribute, attRef.MTextAttribute.Contents, attRef.TextString))
                attRef.Erase()
            End If
        Next
        For Each attDef As AttributeDefinition In attDefs
            Dim attRef As New AttributeReference()
            attRef.SetAttributeFromBlock(attDef, br.BlockTransform)
            If attValues IsNot Nothing AndAlso attValues.ContainsKey(attDef.Tag) Then
                attRef.TextString = attValues(attDef.Tag.ToUpper())
            End If
            br.AttributeCollection.AppendAttribute(attRef)
            tm.AddNewlyCreatedDBObject(attRef, True)
        Next
    End Sub

    Public Function Creaza_Mleader_nou_fara_UCS_transform(ByVal Point1 As Point3d, ByVal Continut As String, ByVal text_height As Double, ByVal arrow_size As Double, ByVal Gap1 As Double, ByVal DELTA_X As Double, ByVal DELTA_Y As Double) As MLeader
        Try
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim Mtext1 As New MText
                Mtext1.TextHeight = text_height
                Mtext1.Contents = "{\H" & text_height.ToString & ";" & Continut & "}"
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


    Public Function Creaza_polyline_din_polyline3d(ByVal Trans1 As Transaction, ByVal PolyCL3D As Polyline3d)

        If IsNothing(PolyCL3D) = False Then
            Dim Index_Poly As Integer = 0
            Dim PolyCL As New Polyline
            For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In PolyCL3D
                Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                        (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                Dim x1 As Double = v3d.Position.X
                Dim y1 As Double = v3d.Position.Y
                Dim z1 As Double = v3d.Position.Z
                PolyCL.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                Index_Poly = Index_Poly + 1
            Next

            Return PolyCL
        End If

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

    Public Sub Add_to_clipboard_2_Data_table(ByVal Data_table1 As System.Data.DataTable, ByVal Data_table2 As System.Data.DataTable)
        Dim sTR1 As String = ""
        If Data_table1.Rows.Count > 0 Then
            For i = 0 To Data_table1.Columns.Count - 1
                If i = 0 Then
                    sTR1 = Data_table1.Columns(i).ColumnName
                Else
                    sTR1 = sTR1 & Chr(9) & Data_table1.Columns(i).ColumnName
                End If
            Next
            For i = 0 To Data_table1.Rows.Count - 1
                sTR1 = sTR1 & vbCrLf
                For j = 0 To Data_table1.Columns.Count - 1
                    If IsDBNull(Data_table1.Rows(i).Item(j)) = False Then
                        If j = 0 Then
                            sTR1 = sTR1 & Data_table1.Rows(i).Item(j)
                        Else
                            sTR1 = sTR1 & Chr(9) & Data_table1.Rows(i).Item(j)
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

        If IsNothing(Data_table2) = False Then
            If Data_table2.Rows.Count > 0 Then
                sTR1 = sTR1 & vbCrLf
                For i = 0 To Data_table2.Columns.Count - 1
                    If i = 0 Then
                        sTR1 = sTR1 & Data_table2.Columns(i).ColumnName
                    Else
                        sTR1 = sTR1 & Chr(9) & Data_table2.Columns(i).ColumnName
                    End If
                Next
                For i = 0 To Data_table2.Rows.Count - 1
                    sTR1 = sTR1 & vbCrLf
                    For j = 0 To Data_table2.Columns.Count - 1
                        If IsDBNull(Data_table2.Rows(i).Item(j)) = False Then
                            If j = 0 Then
                                sTR1 = sTR1 & Data_table2.Rows(i).Item(j)
                            Else
                                sTR1 = sTR1 & Chr(9) & Data_table2.Rows(i).Item(j)
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
        End If
        My.Computer.Clipboard.SetText(sTR1)
    End Sub
    Public Sub add_extra_param_to_dim(ByVal dimension1 As RotatedDimension, ByVal thisdrawing As Document)
        With dimension1
            .Dimpost = "<>" 'prefix
            .Dimrnd = 0 'Rounds all dimensioning distances to the specified value. 

            .Dimtxtdirection = False
            'Specifies the reading direction of the dimension text. 
            .Dimtofl = False
            'Initial value: Off (imperial) or On (metric)  
            .Dimtoh = False
            'Controls the position of dimension text outside the extension lines. 
            .Dimtih = False
            'Initial value: On (imperial) or Off (metric)  
            .Dimtad = 0
            'Controls the vertical position of text in relation to the dimension line. 
            .Dimtvp = 0
            'Controls the vertical position of dimension text above or below the dimension line. 
            .Dimsd1 = False
            'Controls suppression of the first dimension line and arrowhead. 
            .Dimsd2 = False
            'Controls suppression of the second dimension line and arrowhead. 
            .Dimse1 = True 'Suppresses display of the first extension line. 
            .Dimse2 = True 'Suppresses display of the second extension line
            .Dimjust = 0
            'Controls the horizontal positioning of dimension text. 
            .Dimadec = 0 'Controls the number of precision places displayed in angular dimensions. (0-8)
            .Dimalt = False 'Controls the display of alternate units in dimensions. Off - Disables alternate units
            .Dimaltd = 2 'Controls the number of decimal places in alternate units. If DIMALT is turned on, DIMALTD sets the number of digits displayed to the right of the decimal point in the alternate measurement
            .Dimaltf = 25.4 'Controls the multiplier for alternate units. If DIMALT is turned on, DIMALTF multiplies linear dimensions by a factor to produce a value in an alternate system of measurement. The initial value represents the number of millimeters in an inch.
            .Dimaltmzf = 100
            .Dimaltrnd = 0 'Rounds off the alternate dimension units. 
            .Dimalttd = 2 'Sets the number of decimal places for the tolerance values in the alternate units of a dimension. 
            .Dimalttz = 0 'Controls suppression of zeros in tolerance values. 
            .Dimaltu = 2 'Sets the units format for alternate units of all dimension substyles except Angular. (2 - Decimal)
            .Dimaltz = 0 'Controls the suppression of zeros for alternate unit dimension values. 
            .Dimapost = "" 'Specifies a text prefix or suffix (or both) to the alternate dimension measurement for all types of dimensions except angular. 
            .Dimarcsym = 0 'Controls display of the arc symbol in an arc length dimension. (0- Places arc length symbols before the dimension text )
            .Dimatfit = 3
            'Determines how dimension text and arrows are arranged when space is not sufficient to place both within the extension lines. 
            .Dimaunit = 0 'Sets the units format for angular dimensions. (0 - Decimal degrees)
            .Dimazin = 0 'Suppresses zeros for angular dimensions. 
            .Dimsah = False
            'Controls the display of dimension line arrowhead blocks. 
            .Dimcen = 0.09 'Controls drawing of circle or arc center marks and centerlines by the DIMCENTER, DIMDIAMETER, and DIMRADIUS commands. 
            .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) ' Assigns colors to dimension lines, arrowheads, and dimension leader lines
            .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension extension lines.
            .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension text
            .Dimdle = 0 'Sets the distance the dimension line extends beyond the extension line when oblique strokes are drawn instead of arrowheads. 
            .Dimdli = 0.38 'Controls the spacing of the dimension lines in baseline dimensions. 
            'Each dimension line is offset from the previous one by this amount, if necessary, to avoid drawing over it. Changes made with DIMDLI are not applied to existing dimensions
            .Dimdsep = ".c"
            'Specifies a single-character decimal separator to use when creating dimensions whose unit format is decimal
            .Dimexe = 0.18 'Specifies how far to extend the extension line beyond the dimension line. 
            .Dimexo = 0.0625 'Specifies how far extension lines are offset from origin points. 
            'With fixed-length extension lines, this value determines the minimum offset. 
            .Dimfrac = 0 'Sets the fraction format when DIMLUNIT is set to 4 (Architectural) or 5 (Fractional).

            .Dimfxlen = 1
            .DimfxlenOn = False
            .Dimgap = 0.09 'Sets the distance around the dimension text when the dimension line breaks to accommodate dimension text.
            .Dimjogang = 0.785398163 'Determines the angle of the transverse segment of the dimension line in a jogged radius dimension. 
            .Dimlfac = 1
            'Sets a scale factor for linear dimension measurements. 
            .Dimltex1 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the first extension line. 
            .Dimltex2 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the second extension line. 
            .Dimltype = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the dimension line.
            .Dimlunit = 2
            'Sets units for all dimension types except Angular. 
            .Dimlwd = LineWeight.ByBlock
            'Assigns lineweight to dimension lines. 
            .Dimlwe = LineWeight.ByBlock
            'Assigns lineweight to extension  lines. 
            .Dimmzf = 100
            .Dimscale = 1
            'Sets the overall scale factor applied to dimensioning variables that specify sizes, distances, or offsets. 
            .Dimtdec = 0
            'Sets the number of decimal places to display in tolerance values for the primary units in a dimension. 
            .Dimtfac = 1
            'Specifies a scale factor for the text height of fractions and tolerance values relative to the dimension text height, as set by DIMTXT. 
            .Dimtfill = 1
            'Controls the background of dimension text. 
            .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)
            .Dimtix = False
            'Draws text between extension lines. 
            .Dimsoxd = False
            'Suppresses arrowheads if not enough space is available inside the extension lines. 
            .Dimtm = 0
            'Sets the minimum (or lower) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            .Dimtmove = 0
            'Sets dimension text movement rules. 
            .Dimtp = 0
            'Sets the maximum (or upper) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            .Dimlim = False
            'Generates dimension limits as the default text. 
            .Dimtol = False
            'Appends tolerances to dimension text. 
            .Dimtolj = 1 'Sets the vertical justification for tolerance values relative to the nominal dimension text. 
            .Dimtsz = 0
            'Specifies the size of oblique strokes drawn instead of arrowheads for linear, radius, and diameter dimensioning. 
            .Dimtzin = 0 'Controls the suppression of zeros in tolerance values. 
            .Dimupt = False
            'Controls options for user-positioned text. 
            .Dimzin = 0
            'Controls the suppression of zeros in the primary unit value. 




        End With


    End Sub

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
