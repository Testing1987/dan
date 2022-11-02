Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

'Editor1.GetCurrentView.ViewDirection


Public Class Commands_class_plain_autocad
    Public Shared Profile_table As New System.Data.DataTable
    Public Shared Total_chainage_grid As Double
    Public Shared Poly3Did As ObjectId
    Dim RowXL As Integer = 1
    <CommandMethod("CPU")>
    Public Sub gaseste_nr_serial_al_disc_C()
        Dim disk As New Management.ManagementObject("Win32_LogicalDisk.DeviceID=""C:""")
        Dim diskPropertyB As Management.PropertyData = disk.Properties("VolumeSerialNumber")
        MsgBox(diskPropertyB.Value.ToString())
    End Sub

    <CommandMethod("read_xData")>
    Public Sub read_xData()
        If isSECURE() = False Then Exit Sub
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Editor1 = ThisDrawing.Editor
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select obiect:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If


            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead)
                Dim REG_APP_TABLE As RegAppTable = Trans1.GetObject(ThisDrawing.Database.RegAppTableId, OpenMode.ForRead)

                Dim DataTable1 As New System.Data.DataTable
                DataTable1.Columns.Add("OBJECT_ID", GetType(String))
                DataTable1.Columns.Add("INDEX", GetType(Integer))
                DataTable1.Columns.Add("APP", GetType(String))
                DataTable1.Columns.Add("CODE", GetType(String))
                DataTable1.Columns.Add("VALUE", GetType(String))


                Dim Colectie1 As New Specialized.StringCollection

                For Each id1 As ObjectId In REG_APP_TABLE
                    Dim reg_app_table_record As RegAppTableRecord = Trans1.GetObject(id1, OpenMode.ForRead)
                    Colectie1.Add(reg_app_table_record.Name)
                Next



                For i = 1 To Rezultat1.Value.Count



                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                    Obj1 = Rezultat1.Value.Item(i - 1)

                    Dim Ent1 As Entity
                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                    Dim Steel_counter As Integer = 0
                    Dim String1 As String


                    For j = 0 To Colectie1.Count - 1


                        Dim ResultBuffer_1 As ResultBuffer = Ent1.GetXDataForApplication(Colectie1(j))
                        If IsNothing(ResultBuffer_1) = False Then
                            Dim Index As Integer = 0

                            For Each TypedValue1 As TypedValue In ResultBuffer_1
                                DataTable1.Rows.Add()
                                DataTable1.Rows(DataTable1.Rows.Count - 1).Item("OBJECT_ID") = Ent1.ObjectId.Handle.Value.ToString
                                DataTable1.Rows(DataTable1.Rows.Count - 1).Item("APP") = Colectie1(j)
                                DataTable1.Rows(DataTable1.Rows.Count - 1).Item("INDEX") = Index
                                DataTable1.Rows(DataTable1.Rows.Count - 1).Item("CODE") = TypedValue1.TypeCode
                                DataTable1.Rows(DataTable1.Rows.Count - 1).Item("VALUE") = TypedValue1.Value.ToString
                                Index = Index + 1
                            Next
                        End If
                    Next

                    Functions.Transfer_datatable_to_new_excel_spreadsheet(DataTable1)


                Next
            End Using ' asta e de la trans1
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Catch ex As Exception

            MsgBox(ex.Message)
        End Try


    End Sub


    <CommandMethod("PPLWORKSPACE")>
    Public Sub Show_WorkSpace_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is WORKSPACE_FORM Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New WORKSPACE_FORM
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("Clone_and_replace_MTEXT", CommandFlags.UsePickSet)>
    Public Sub Clone_and_replace_object()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)




                        For i = 0 To Rezultat1.Value.Count - 1

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForWrite)
                            If TypeOf Ent1 Is MText Then
                                Dim Mtext1 As MText = Ent1
                                Dim Ent2 As New MText
                                Ent2.Contents = Mtext1.Contents
                                Ent2.Layer = Mtext1.Layer
                                Ent2.Rotation = Mtext1.Rotation
                                'Ent2.TextStyleId = Mtext1.TextStyleId
                                Ent2.TextHeight = Mtext1.TextHeight
                                Ent2.ColorIndex = Mtext1.ColorIndex
                                Ent2.LineWeight = Mtext1.LineWeight
                                Ent2.Location = Mtext1.Location
                                Ent2.Attachment = Mtext1.Attachment
                                Ent2.Width = Mtext1.Width
                                BTrecord.AppendEntity(Ent2)
                                Trans1.AddNewlyCreatedDBObject(Ent2, True)

                            End If





                            Ent1.Erase()




                        Next


                        Trans1.Commit()
                        Editor1.Regen()


                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("ppldrotated")>
    Public Sub dim_linear_rotated()
        If isSECURE() = False Then Exit Sub

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Try

            'ascunde_butoanele_pentru_forms(Me, colectie1)
            Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
            Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify first point:")
            PP1.AllowNone = False
            Point1 = Editor1.GetPoint(PP1)

            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                'afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub
            End If

            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify second point:")
            PP2.UseBasePoint = True
            PP2.BasePoint = Point1.Value

            PP2.AllowNone = False
            Point2 = Editor1.GetPoint(PP2)

            If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                'afiseaza_butoanele_pentru_forms(Me, colectie1)
                Exit Sub
            End If



            Using Lock_dwg As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Dimension1 As New RotatedDimension
                    Dim Rotatie As Double = GET_Bearing_rad(Point1.Value.TransformBy(CurentUCSmatrix).X, Point1.Value.TransformBy(CurentUCSmatrix).Y, Point2.Value.TransformBy(CurentUCSmatrix).X, Point2.Value.TransformBy(CurentUCSmatrix).Y)

                    Dimension1.XLine1Point = Point1.Value.TransformBy(CurentUCSmatrix)
                    Dimension1.XLine2Point = Point2.Value.TransformBy(CurentUCSmatrix)
                    Dimension1.Rotation = Rotatie
                    Dimension1.DimLinePoint = Point2.Value.TransformBy(CurentUCSmatrix)
                    Dimension1.UsingDefaultTextPosition = True
                    Dimension1.TextAttachment = AttachmentPoint.MiddleCenter

                    'Dimension1.TextPosition = Int1(0)


                    Dimension1.TextRotation = Rotatie

                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)


                    Dim Exista As Boolean = False
                    Dim Text_style_romans As TextStyleTableRecord

                    For Each TextStyle_id As ObjectId In Text_style_table
                        Dim TextStyle As TextStyleTableRecord = Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        With TextStyle
                            If .FileName = "romans.shx" And .XScale = 1.0 And .ObliquingAngle = 0 Then
                                Exista = True
                                Text_style_romans = TextStyle
                                Exit For
                            End If
                        End With
                    Next


                    If Exista = False Then
                        Text_style_table.UpgradeOpen()
                        Text_style_romans = New TextStyleTableRecord
                        Text_style_romans.Name = "ROMANS"

                        Text_style_romans.TextSize = 0
                        Text_style_romans.ObliquingAngle = 0
                        Text_style_romans.FileName = "romans.shx"
                        Text_style_romans.XScale = 1.0
                        Text_style_table.Add(Text_style_romans)
                        Trans1.AddNewlyCreatedDBObject(Text_style_romans, True)

                    End If

                    Dim Arrowid As ObjectId
                    With Dimension1
                        .Dimasz = 18 'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                        'Multiples of the arrowhead size determine whether dimension lines and text should fit between the extension lines. DIMASZ is also used to scale arrowhead blocks if set by DIMBLK. DIMASZ has no effect when DIMTSZ is other than zero

                        .Dimdec = 0
                        'Sets the number of decimal places displayed for the primary units of a dimension
                        'The precision is based on the units or angle format you have selected. 


                        .Dimtxt = 8 'Specifies the height of dimension text, unless the current text style has a fixed height

                        .TextStyleId = Text_style_romans.ObjectId

                        .Dimtxtdirection = False
                        'Specifies the reading direction of the dimension text. 
                        '0 - Displays dimension text in a Left-to-Right reading style 
                        '1 - Displays dimension text in a Right-to-Left reading style  



                        .Dimtofl = False
                        'Initial value: Off (imperial) or On (metric)  
                        'Controls whether a dimension line is drawn between the extension lines even when the text is placed outside. 
                        'For radius and diameter dimensions (when DIMTIX is off), draws a dimension line inside the circle or arc and places the text, arrowheads, and leader outside. 
                        ' Off -  Does not draw dimension lines between the measured points when arrowheads are placed outside the measured points 
                        ' On -  Draws dimension lines between the measured points even when arrowheads are placed outside the measured points 

                        .Dimtoh = False
                        'Controls the position of dimension text outside the extension lines. 
                        ' Off -  Aligns text with the dimension line
                        ' On -  Draws text horizontally

                        .Dimtih = False
                        'Initial value: On (imperial) or Off (metric)  
                        'Controls the position of dimension text inside the extension lines for all dimension types except Ordinate. 
                        'Off - Aligns text with the dimension line
                        'On -  Draws text horizontally

                        .Dimtad = 0
                        'Controls the vertical position of text in relation to the dimension line. 
                        '0 - Centers the dimension text between the extension lines. 
                        '1 - Places the dimension text above the dimension line except when the dimension line is not horizontal and text inside the extension lines is forced horizontal ( DIMTIH = 1). 
                        '    The distance from the dimension line to the baseline of the lowest line of text is the current DIMGAP value. 
                        '2 - Places the dimension text on the side of the dimension line farthest away from the defining points. 
                        '3 - Places the dimension text to conform to Japanese Industrial Standards (JIS). 
                        '4 - Places the dimension text below the dimension line. 


                        .Dimtvp = 0
                        'Controls the vertical position of dimension text above or below the dimension line. 
                        'The DIMTVP value is used when DIMTAD is off. The magnitude of the vertical offset of text is the product of the text height and DIMTVP. 
                        'Setting DIMTVP to 1.0 is equivalent to setting DIMTAD to on. The dimension line splits to accommodate the text only if the absolute value of DIMTVP is less than 0.7. 


                        .Dimsd1 = False
                        'Controls suppression of the first dimension line and arrowhead. 
                        'When turned on, suppresses the display of the dimension line and arrowhead between the first extension line and the text. 
                        .Dimsd2 = False
                        'Controls suppression of the second dimension line and arrowhead. 
                        'When turned on, suppresses the display of the dimension line and arrowhead between the second extension line and the text. 
                        .Dimse1 = True 'Suppresses display of the first extension line. 
                        .Dimse2 = True 'Suppresses display of the second extension line

                        .Dimrnd = 5
                        'Rounds all dimensioning distances to the specified value. 
                        'For instance, if DIMRND is set to 0.25, all distances round to the nearest 0.25 unit. 
                        'If you set DIMRND to 1.0, all distances round to the nearest integer. 
                        'Note that the number of digits edited after the decimal point depends on the precision set by DIMDEC. DIMRND does not apply to angular dimensions. 

                        .Dimpost = "<>'"
                        'Specifies a text prefix or suffix (or both) to the dimension measurement. 
                        'For example, to establish a suffix for millimeters, set DIMPOST to mm; a distance of 19.2 units would be displayed as 19.2 mm. 
                        'If tolerances are turned on, the suffix is applied to the tolerances as well as to the main dimension. 
                        'Use <> to indicate placement of the text in relation to the dimension value. 
                        'For example, enter <>mm to display a 5.0 millimeter radial dimension as "5.0mm." 
                        'If you entered mm <>, the dimension would be displayed as "mm 5.0." 
                        'Use the <> mechanism for angular dimensions. 

                        .Dimjust = 0
                        'Controls the horizontal positioning of dimension text. 
                        '0 -  Positions the text above the dimension line and center-justifies it between the extension lines 
                        '1 -  Positions the text next to the first extension line 
                        '2 -  Positions the text next to the second extension line 
                        '3 -  Positions the text above and aligned with the first extension line 
                        '4 -  Positions the text above and aligned with the second extension line 

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
                        'For instance, if the current units are Architectural, DIMALT is on, DIMALTF is 25.4 (the number of millimeters per inch), DIMALTD is 2, and DIMPOST is set to "mm," a distance of 10 units would be displayed as 10"[254.00mm]. 
                        'To turn off an established prefix or suffix (or both), set it to a single period (.). 
                        .Dimarcsym = 0 'Controls display of the arc symbol in an arc length dimension. (0- Places arc length symbols before the dimension text )
                        '1 - Places arc length symbols above the dimension text 
                        '2 -  Suppresses the display of arc length symbols 

                        .Dimatfit = 3
                        'Determines how dimension text and arrows are arranged when space is not sufficient to place both within the extension lines. 
                        '0 -  Places both text and arrows outside extension lines 
                        '1 -  Moves arrows first, then text
                        '2 -  Moves text first, then arrows
                        '3 -  Moves either text or arrows, whichever fits best 
                        'A leader is added to moved dimension text when DIMTMOVE is set to 1. 


                        .Dimaunit = 0 'Sets the units format for angular dimensions. (0 - Decimal degrees)
                        .Dimazin = 0 'Suppresses zeros for angular dimensions. 


                        .Dimsah = False
                        'Controls the display of dimension line arrowhead blocks. 
                        'Off - Use arrowhead blocks set by DIMBLK
                        'On - Use arrowhead blocks set by DIMBLK1 and DIMBLK2

                        .Dimblk = Arrowid
                        'Sets the arrowhead block displayed at the ends of dimension lines or leader lines. 
                        'To return to the default, closed-filled arrowhead display, enter a single period (.). Arrowhead block entries and the names used to select them in the New, Modify, and Override Dimension Style dialog boxes are shown below. You can also enter the names of user-defined arrowhead blocks. 
                        'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
                        '"" - Closed(filled)
                        '"_DOT" - dot
                        '"_DOTSMALL" - dot small
                        '"_DOTBLANK" - dot blank
                        '"_ORIGIN" - origin indicator
                        '"_ORIGIN2" - origin indicator 2
                        '"_OPEN" - open
                        '"_OPEN90" - Right(angle)
                        '"_OPEN30" - open 30
                        '"_CLOSED" - Closed
                        '"_SMALL" - dot small blank
                        '"_NONE" - none
                        '"_OBLIQUE" - oblique
                        '"_BOXFILLED" - box filled
                        '"_BOXBLANK" - box
                        '"_CLOSEDBLANK" - Closed(blank)
                        '"_DATUMFILLED" - datum triangle filled
                        '"_DATUMBLANK" - datum triangle
                        '"_INTEGRAL" - integral
                        '"_ARCHTICK" - architectural tick


                        .Dimblk1 = Arrowid
                        'Sets the arrowhead for the first end of the dimension line when DIMSAH is on. 
                        'To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowheads, see DIMBLK. 
                        'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
                        .Dimblk2 = Arrowid
                        'Sets the arrowhead for the second end of the dimension line when DIMSAH is on. 
                        'To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowhead entries, see DIMBLK. 
                        'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
                        .Dimldrblk = Arrowid ' Specifies the arrow type for leaders. 

                        .Dimcen = 0.09 'Controls drawing of circle or arc center marks and centerlines by the DIMCENTER, DIMDIAMETER, and DIMRADIUS commands. 
                        .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) ' Assigns colors to dimension lines, arrowheads, and dimension leader lines
                        .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension extension lines.
                        .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension text


                        .Dimdle = 0 'Sets the distance the dimension line extends beyond the extension line when oblique strokes are drawn instead of arrowheads. 
                        .Dimdli = 0.38 'Controls the spacing of the dimension lines in baseline dimensions. 
                        'Each dimension line is offset from the previous one by this amount, if necessary, to avoid drawing over it. Changes made with DIMDLI are not applied to existing dimensions
                        .Dimdsep = ".c"
                        'Specifies a single-character decimal separator to use when creating dimensions whose unit format is decimal
                        'When prompted, enter a single character at the Command prompt. If dimension units is set to Decimal, the DIMDSEP character is used instead of the default decimal point.
                        'If DIMDSEP is set to NULL (default value, reset by entering a period), the decimal point is used as the dimension separator
                        .Dimexe = 0.18 'Specifies how far to extend the extension line beyond the dimension line. 
                        .Dimexo = 0.0625 'Specifies how far extension lines are offset from origin points. 
                        'With fixed-length extension lines, this value determines the minimum offset. 
                        .Dimfrac = 0 'Sets the fraction format when DIMLUNIT is set to 4 (Architectural) or 5 (Fractional).
                        '0 - Horizontal stacking
                        '1 - Diagonal stacking
                        '2 - Not stacked (for example, 1/2)


                        .Dimfxlen = 1
                        .DimfxlenOn = False

                        .Dimgap = 0.09 'Sets the distance around the dimension text when the dimension line breaks to accommodate dimension text.
                        .Dimjogang = 0.785398163 'Determines the angle of the transverse segment of the dimension line in a jogged radius dimension. 



                        .Dimlfac = 1
                        'Sets a scale factor for linear dimension measurements. 
                        'All linear dimension distances, including radii, diameters, and coordinates, are multiplied by DIMLFAC before being converted to dimension text. Positive values of DIMLFAC are applied to dimensions in both model space and paper space; negative values are applied to paper space only. 
                        'DIMLFAC applies primarily to nonassociative dimensions (DIMASSOC set 0 or 1). For nonassociative dimensions in paper space, DIMLFAC must be set individually for each layout viewport to accommodate viewport scaling. 
                        'DIMLFAC has no effect on angular dimensions, and is not applied to the values held in DIMRND, DIMTM, or DIMTP. 

                        .Dimltex1 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the first extension line. 
                        .Dimltex2 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the second extension line. 
                        .Dimltype = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the dimension line.

                        .Dimlunit = 2
                        'Sets units for all dimension types except Angular. 
                        '1 Scientific
                        '2 Decimal
                        '3 Engineering
                        '4 Architectural (always displayed stacked)
                        '5 Fractional (always displayed stacked)
                        '6 Microsoft Windows Desktop (decimal format using Control Panel settings for decimal separator and number grouping symbols) 


                        .Dimlwd = LineWeight.ByBlock
                        'Assigns lineweight to dimension lines. 
                        '-3 Default (the LWDEFAULT value) 
                        '-2 BYBLOCK
                        '-1 BYLAYER

                        .Dimlwe = LineWeight.ByBlock
                        'Assigns lineweight to extension  lines. 
                        '-3 Default (the LWDEFAULT value) 
                        '-2 BYBLOCK
                        '-1 BYLAYER



                        .Dimmzf = 100


                        .Dimscale = 1
                        'Sets the overall scale factor applied to dimensioning variables that specify sizes, distances, or offsets. 
                        'Also affects the leader objects with the LEADER command. 
                        'Use MLEADERSCALE to scale multileader objects created with the MLEADER command. 
                        '0.0 - A reasonable default value is computed based on the scaling between the current model space viewport and paper space. 
                        'If you are in paper space or model space and not using the paper space feature, the scale factor is 1.0. 
                        '>0 - A scale factor is computed that leads text sizes, arrowhead sizes, and other scaled distances to plot at their face values. 
                        'DIMSCALE does not affect measured lengths, coordinates, or angles. 
                        'Use DIMSCALE to control the overall scale of dimensions. However, if the current dimension style is annotative, 
                        'DIMSCALE is automatically set to zero and the dimension scale is controlled by the CANNOSCALE system variable. DIMSCALE cannot be set to a non-zero value when using annotative dimensions. 

                        .Dimtdec = 0
                        'Sets the number of decimal places to display in tolerance values for the primary units in a dimension. 
                        'This system variable has no effect unless DIMTOL is set to On. The default for DIMTOL is Off. 

                        .Dimtfac = 1
                        'Specifies a scale factor for the text height of fractions and tolerance values relative to the dimension text height, as set by DIMTXT. 
                        'For example, if DIMTFAC is set to 1.0, the text height of fractions and tolerances is the same height as the dimension text. 
                        'If DIMTFAC is set to 0.7500, the text height of fractions and tolerances is three-quarters the size of dimension text. 
                        .Dimtfill = 1
                        'Controls the background of dimension text. 
                        '0 -  No Background
                        '1 -  The background color of the drawing 
                        '2 -  The background specified by DIMTFILLCLR
                        .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)

                        .Dimtix = False
                        'Draws text between extension lines. 
                        'Off -  Varies with the type of dimension. 
                        '        For linear and angular dimensions, text is placed inside the extension lines if there is sufficient room. 
                        '        For radius and diameter dimensions that don't fit inside the circle or arc, DIMTIX has no effect and always forces the text outside the circle or arc. 
                        'On -  Draws dimension text between the extension lines even if it would ordinarily be placed outside those lines 

                        .Dimsoxd = False
                        'Suppresses arrowheads if not enough space is available inside the extension lines. 
                        'Off -  Arrowheads are not suppressed
                        'On -  Arrowheads are suppressed
                        'If not enough space is available inside the extension lines and DIMTIX is on, setting DIMSOXD to On suppresses the arrowheads. If DIMTIX is off, DIMSOXD has no effect. 


                        .Dimtm = 0
                        'Sets the minimum (or lower) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
                        'DIMTM accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
                        'If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower, and a plus sign is added to the DIMTP value if it is positive. 
                        'For DIMTM, the program uses the negative of the value you enter (adding a minus sign if you specify a positive number and a plus sign if you specify a negative number). 

                        .Dimtmove = 0
                        'Sets dimension text movement rules. 
                        '0 -  Moves the dimension line with dimension text
                        '1 -  Adds a leader when dimension text is moved
                        '2 -  Allows text to be moved freely without a leader

                        .Dimtp = 0
                        'Sets the maximum (or upper) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
                        'DIMTP accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
                        'If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower and a plus sign is added to the DIMTP value if it is positive. 


                        .Dimlim = False
                        'Generates dimension limits as the default text. 
                        'Setting DIMLIM to On turns DIMTOL off. 
                        'Off -  Dimension limits are not generated as default text 
                        'On -  Dimension limits are generated as default text


                        .Dimtol = False
                        'Appends tolerances to dimension text. 
                        'Setting DIMTOL to on turns DIMLIM off. 

                        .Dimtolj = 1 'Sets the vertical justification for tolerance values relative to the nominal dimension text. 



                        .Dimtsz = 0
                        'Specifies the size of oblique strokes drawn instead of arrowheads for linear, radius, and diameter dimensioning. 
                        '0 -  Draws arrowheads.
                        '>0 -  Draws oblique strokes instead of arrowheads. The size of the oblique strokes is determined by this value multiplied by the DIMSCALE value 




                        .Dimtzin = 0 'Controls the suppression of zeros in tolerance values. 

                        .Dimupt = False
                        'Controls options for user-positioned text. 
                        'Off -  Cursor controls only the dimension line location
                        'On -  Cursor controls both the text position and the dimension line location 

                        .Dimzin = 0
                        'Controls the suppression of zeros in the primary unit value. 
                        'Values 0-3 affect feet-and-inch dimensions only: 
                        '0 -  Suppresses zero feet and precisely zero inches
                        '1 -  Includes zero feet and precisely zero inches
                        '2 -  Includes zero feet and suppresses zero inches
                        '3 -  Includes zero inches and suppresses zero feet
                        '4 -  Suppresses leading zeros in decimal dimensions (for example, 0.5000 becomes .5000) 
                        '8 -  Suppresses trailing zeros in decimal dimensions (for example, 12.5000 becomes 12.5) 
                        '12 -  Suppresses both leading and trailing zeros (for example, 0.5000 becomes .5) 



                    End With




                    BTrecord.AppendEntity(Dimension1)
                    Trans1.AddNewlyCreatedDBObject(Dimension1, True)
                    Trans1.Commit()
                End Using
            End Using


            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("pplprof", CommandFlags.UsePickSet)>
    Public Sub Creaza_graph_profile_from_a_3d_polyline()
        If isSECURE() = False Then Exit Sub
        Total_chainage_grid = 0
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document
        ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Dim CSF As Double = 1
        Try
            Dim Ground_3d_length As Double = 0




            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select 3dpolyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If

            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                Dim Pline3 As Autodesk.AutoCAD.DatabaseServices.Polyline3d


                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj1 = Rezultat1.Value.Item(0)
                Dim Ent1 As Entity
                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                    If Profile_table.Columns.Contains("X") = False Then Profile_table.Columns.Add("X", GetType(Double))
                    If Profile_table.Columns.Contains("Y") = False Then Profile_table.Columns.Add("Y", GetType(Double))
                    If Profile_table.Columns.Contains("Z") = False Then Profile_table.Columns.Add("Z", GetType(Double))
                    If Profile_table.Columns.Contains("Partial_Chainage_grid") = False Then Profile_table.Columns.Add("Partial_Chainage_grid", GetType(Double))
                    If Profile_table.Columns.Contains("Total_Chainage_grid") = False Then Profile_table.Columns.Add("Total_Chainage_grid", GetType(Double))
                    If Profile_table.Columns.Contains("Modified_Chainage_grid") = False Then Profile_table.Columns.Add("Modified_Chainage_grid", GetType(Double))

                    Poly3Did = Obj1.ObjectId
                    Ent1.UpgradeOpen()

                    Pline3 = Ent1
                    Dim Min_elev, Max_elev As Double

                    Dim i As Double = 0

                    For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In Pline3
                        Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                        Dim x1 As Double = v3d.Position.X
                        Dim y1 As Double = v3d.Position.Y
                        Dim z1 As Double = v3d.Position.Z
                        'Public Shared Prof_X As New DoubleCollection
                        'Public Shared Prof_Y As New DoubleCollection
                        'Public Shared Prof_Z As New DoubleCollection
                        'Prof_X.Add(x1)
                        'Prof_Y.Add(y1)
                        'Prof_Z.Add(z1)

                        Profile_table.Rows.Add()
                        Profile_table.Rows.Item(i).Item("X") = x1
                        Profile_table.Rows.Item(i).Item("Y") = y1
                        Profile_table.Rows.Item(i).Item("Z") = z1
                        Profile_table.Rows.Item(i).Item("Modified_Chainage_grid") = 0


                        If i = 0 Then
                            Min_elev = z1
                            Max_elev = z1
                            Profile_table.Rows.Item(i).Item("Partial_Chainage_grid") = 0
                            Profile_table.Rows.Item(i).Item("Total_Chainage_grid") = 0
                        Else
                            Total_chainage_grid = Total_chainage_grid + GET_distanta_Double_XY(Profile_table.Rows.Item(i - 1).Item("X"), Profile_table.Rows.Item(i - 1).Item("Y"), x1, y1)
                            Profile_table.Rows.Item(i).Item("Total_Chainage_grid") = Total_chainage_grid
                            Profile_table.Rows.Item(i).Item("Partial_Chainage_grid") = GET_distanta_Double_XY(Profile_table.Rows.Item(i - 1).Item("X"), Profile_table.Rows.Item(i - 1).Item("Y"), x1, y1)
                            Ground_3d_length = Ground_3d_length + GET_distanta3d_Double_with_CSF(Profile_table.Rows.Item(i - 1).Item("X"), Profile_table.Rows.Item(i - 1).Item("Y"), Profile_table.Rows.Item(i - 1).Item("Z"), x1, y1, z1, CSF)

                        End If




                        If z1 < Min_elev Then Min_elev = z1
                        If z1 > Max_elev Then Max_elev = z1
                        i = i + 1
                    Next



                    For Each forma In System.Windows.Forms.Application.OpenForms

                        If TypeOf forma Is profiler3d_form Then
                            forma.Focus()
                            forma.WindowState = Windows.Forms.FormWindowState.Normal
                            Exit Sub
                        End If
                    Next


                    Try
                        Dim forma1 As New profiler3d_form
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
                        forma1.TextBox_L_elevation.Text = Fix(Min_elev) - 1
                        forma1.TextBox_H_Elevation.Text = Fix(Max_elev) + 1
                        forma1.TextBox_Ground_Chainage.Text = (Total_chainage_grid / CSF).ToString
                        forma1.TextBox_3d_length.Text = Ground_3d_length.ToString
                        Incarca_existing_LINETYPES_to_combobox(forma1.ComboBox_LINETYPE)
                        Incarca_existing_textstyles_to_combobox(forma1.ComboBox_text_styles)
                        Incarca_existing_layers_to_combobox(forma1.ComboBox_LAYER_GRIDLINES)
                        Incarca_existing_layers_to_combobox(forma1.ComboBox_LAYER_POLYLINE)
                        Incarca_existing_layers_to_combobox(forma1.ComboBox_LAYER_TEXT)

                        forma1.TextBox_color_index_grid_lines.Text = "8"
                        forma1.TextBox_text_height.Text = "2.5"
                        If forma1.ComboBox_text_styles.Items.Contains("ROMANS") = True Then
                            forma1.ComboBox_text_styles.Text = "ROMANS"
                        End If
                        If forma1.ComboBox_LINETYPE.Items.Contains("TCHIDDEN") = True Then
                            forma1.ComboBox_LINETYPE.Text = "TCHIDDEN"
                        End If
                        If forma1.ComboBox_LAYER_GRIDLINES.Items.Contains("PGRID") = True Then
                            forma1.ComboBox_LAYER_GRIDLINES.Text = "PGRID"
                        End If
                        If forma1.ComboBox_LAYER_POLYLINE.Items.Contains("PGRADE") = True Then
                            forma1.ComboBox_LAYER_POLYLINE.Text = "PGRADE"
                        End If
                        If forma1.ComboBox_LAYER_TEXT.Items.Contains("TEXT") = True Then
                            forma1.ComboBox_LAYER_TEXT.Text = "TEXT"
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try



                End If ' asta e de la If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d

            End Using

            Editor1.WriteMessage(vbLf & "Command:")
            Editor1.SetImpliedSelection(Empty_array)
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
            Editor1.SetImpliedSelection(Empty_array)
        End Try

    End Sub

    <CommandMethod("PPLFIELDBEND", CommandFlags.UsePickSet)>
    Public Sub Show_Dual_Field_bend_form()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Field_bend_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Field_bend_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPLPCONVERTER")>
    Public Sub Show_profiler_convertor_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Profiler_convertor_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Profiler_convertor_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
            Incarca_existing_LINETYPES_to_combobox(forma1.ComboBox_linetype)
            Incarca_existing_textstyles_to_combobox(forma1.ComboBox_text_style)

            forma1.ComboBox_Vertical_scale_HMM.SelectedIndex = 0



            If forma1.ComboBox_linetype.Items.Contains("TCDOT2") = True Then
                forma1.ComboBox_linetype.Text = "TCDOT2"
            Else
                forma1.ComboBox_linetype.SelectedIndex = 0

            End If
            If forma1.ComboBox_text_style.Items.Contains("ROMANS") = True Then
                forma1.ComboBox_text_style.Text = "ROMANS"
            Else
                forma1.ComboBox_text_style.SelectedIndex = 0

            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPLPROFXL")>
    Public Sub Show_Profiler_from_excel()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Profiler_from_excel_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Profiler_from_excel_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPLKP2PROF")>
    Public Sub Show_chainage_2_autocad_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Chainage_to_graph_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Chainage_to_graph_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PROFCHECK")>
    Public Sub Show_profile_check_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is PROFILE_CHECK_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New PROFILE_CHECK_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPLXLCH2POLY")>
    Public Sub Show_CHAINAGE_TO_POLY_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Chainage_to_polyline_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Chainage_to_polyline_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPL_P_SYMBOL_PICK_PTS")>
    Public Sub INSERT_PIPE_SYMBOL_WITH_PICK_POINTS()

        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Dim Distanta1 As Double

            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Editor1 = ThisDrawing.Editor

            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify the pipe position on ground - source graph:")

            Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify Top of Pipe - source graph")

            Dim PP3 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify the ground position - destination graph:")

            Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem



            Dim Rezultat2 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the Nominal PIPE SIZE in Inches:")
            Rezultat2.AllowNone = False
            Dim Rezultat22 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat2)
            Dim RADIUS_IN_METERS As Double = Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(Rezultat22.Value) / 1000

            Dim Rezultat1 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration - source graph:")
            Rezultat1.AllowNone = True
            Rezultat1.UseDefaultValue = True
            Rezultat1.DefaultValue = 1
            Dim Rezultat11 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat1)
            Dim Exageration_source As Double = Rezultat11.Value
            If Exageration_source <= 0 Then Exageration_source = 1

            Dim Rezultat3 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration - destination graph:")
            Rezultat3.AllowNone = True
            Rezultat3.UseDefaultValue = True
            Rezultat3.DefaultValue = 5
            Dim Rezultat33 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat3)
            Dim Exageration_dest As Double = Rezultat33.Value
            If Exageration_dest <= 0 Then Exageration_dest = 1


            PP1.AllowNone = False
            Point1 = Editor1.GetPoint(PP1)
            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Exit Sub
            End If


            Dim x3, Xinsert, Yinsert As Double
            Dim Y1, Y2, y3 As Double



            Y1 = Point1.Value.TransformBy(UCS_CURENT).Y

            PP2.AllowNone = False
            Point2 = Editor1.GetPoint(PP2)

            If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Exit Sub
            End If



            Y2 = Point2.Value.TransformBy(UCS_CURENT).Y

            PP3.AllowNone = False
            Point3 = Editor1.GetPoint(PP3)

            If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Exit Sub
            End If


            x3 = Point3.Value.TransformBy(UCS_CURENT).X
            y3 = Point3.Value.TransformBy(UCS_CURENT).Y


            Distanta1 = Abs(Y2 - Y1) / Exageration_source

            Xinsert = x3
            Yinsert = y3 - Distanta1 * Exageration_dest - RADIUS_IN_METERS * Exageration_dest
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim layer_curent As LayerTableRecord = Trans1.GetObject(ThisDrawing.Database.Clayer, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim Nume_layer As String = layer_curent.Name
                Dim layerTable As LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If layerTable.Has("TEXT") = True Then
                    Nume_layer = "TEXT"
                End If
                Dim ScaraX, ScaraY As Double
                If Exageration_dest = 1 Then
                    ScaraX = RADIUS_IN_METERS * 2
                    ScaraY = RADIUS_IN_METERS * 2
                Else
                    ScaraY = 2 * RADIUS_IN_METERS * Exageration_dest
                    ScaraX = ScaraY / 2
                End If

                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                'Insereaza_Block_fara_atribute(New Point3d(Xinsert, Yinsert, 0), BTrecord, "YIN_YANG.dwg", "PIPE-YY", Nume_layer, ScaraX, ScaraY, 1, 0)

                Dim colectie_goala As New Specialized.StringCollection
                InsertBlock_with_multiple_atributes("YIN_YANG.dwg", "PIPE-YY", New Point3d(Xinsert, Yinsert, 0), ScaraX, BTrecord, Nume_layer, colectie_goala, colectie_goala)

                Trans1.Commit()
            End Using


            Editor1.WriteMessage(vbLf & "Command:")



        Catch ex As Exception

            Exit Sub
            'MsgBox(ex.Message)
        End Try


    End Sub


    <CommandMethod("PPL_P_SYMBOL_ELEV")>
    Public Sub INSERT_PIPE_SYMBOL_AT_ELEVATION()

        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Dim Distanta1 As Double

            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Editor1 = ThisDrawing.Editor

            Dim Empty_array() As ObjectId


            Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem



            Dim Rezultat2 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the Nominal PIPE SIZE in Inches:")
            Rezultat2.AllowNone = False
            Dim Rezultat22 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat2)
            Dim RADIUS_IN_METERS As Double = Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(Rezultat22.Value) / 1000

            Dim Rezultat_elev_nedeed As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify target elevation:")
            Rezultat_elev_nedeed.AllowNone = True
            Rezultat_elev_nedeed.UseDefaultValue = True
            Rezultat_elev_nedeed.DefaultValue = 0
            Dim Elevatia_dorita As Double = Editor1.GetDouble(Rezultat_elev_nedeed).Value

            Dim Rezultat3 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration:")
            Rezultat3.AllowNone = True
            Rezultat3.UseDefaultValue = True
            Rezultat3.DefaultValue = 1
            Dim Rezultat33 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat3)
            Dim Exageration_dest As Double = Rezultat33.Value
            If Exageration_dest <= 0 Then Exageration_dest = 1

            Dim Rezultat_vert As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_prompt_vert As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_prompt_vert.MessageForAdding = vbLf & "Select a known horizontal line (ELEVATION) and the label for it:"

            Object_prompt_vert.SingleOnly = False
            Rezultat_vert = Editor1.GetSelection(Object_prompt_vert)

            If Not Rezultat_vert.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If



            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify the pipe X position:")
            PP1.AllowNone = False
            Point1 = Editor1.GetPoint(PP1)
            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If





            If Rezultat_vert.Value.Count > 1 Then


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Elevatia_cunoscuta As Double = -100000


                    Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                    Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                    Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line
                    Dim polyLinia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Polyline

                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                    Obj1 = Rezultat_vert.Value.Item(0)
                    Dim Ent1 As Entity
                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                    Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                    Obj2 = Rezultat_vert.Value.Item(1)
                    Dim Ent2 As Entity
                    Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                        mText_cunoscut = Ent1
                        If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                    End If

                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                        mText_cunoscut = Ent2
                        If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                    End If

                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                        Text_cunoscut = Ent1
                        If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                    End If

                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                        Text_cunoscut = Ent2
                        If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                    End If

                    If Elevatia_cunoscuta = -100000 Then

                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If

                    Dim y01, y02 As Double

                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                        Linia_cunoscuta = Ent1
                        y01 = Linia_cunoscuta.StartPoint.Y
                        y02 = Linia_cunoscuta.EndPoint.Y
                        If Abs(y01 - y02) > 0.001 Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End If


                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                        Linia_cunoscuta = Ent2
                        y01 = Linia_cunoscuta.StartPoint.Y
                        y02 = Linia_cunoscuta.EndPoint.Y
                        If Abs(y01 - y02) > 0.001 Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End If

                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        polyLinia_cunoscuta = Ent1
                        y01 = polyLinia_cunoscuta.StartPoint.Y
                        y02 = polyLinia_cunoscuta.EndPoint.Y
                        If Abs(y01 - y02) > 0.001 Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End If

                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        polyLinia_cunoscuta = Ent2
                        y01 = polyLinia_cunoscuta.StartPoint.Y
                        y02 = polyLinia_cunoscuta.EndPoint.Y
                        If Abs(y01 - y02) > 0.001 Then
                            Editor1.SetImpliedSelection(Empty_array)
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End If

                    Dim Semn_plus_minus As Double = 1

                    If Elevatia_dorita < Elevatia_cunoscuta Then
                        Semn_plus_minus = -1
                    End If

                    Distanta1 = Abs(Elevatia_cunoscuta - Elevatia_dorita)


                    Dim Xinsert, Yinsert As Double
                    Xinsert = Point1.Value.TransformBy(UCS_CURENT).X
                    Yinsert = y01 + Semn_plus_minus * Distanta1 * Exageration_dest - RADIUS_IN_METERS * Exageration_dest

                    Dim layer_curent As LayerTableRecord = Trans1.GetObject(ThisDrawing.Database.Clayer, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim Nume_layer As String = layer_curent.Name
                    Dim layerTable As LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    If layerTable.Has("TEXT") = True Then
                        Nume_layer = "TEXT"
                    End If
                    Dim ScaraX, ScaraY As Double
                    If Exageration_dest = 1 Then
                        ScaraX = RADIUS_IN_METERS * 2
                        ScaraY = RADIUS_IN_METERS * 2
                    Else
                        ScaraY = 2 * RADIUS_IN_METERS * Exageration_dest
                        ScaraX = ScaraY / 2
                    End If

                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                    Dim colectie_goala As New Specialized.StringCollection
                    InsertBlock_with_multiple_atributes("YIN_YANG.dwg", "PIPE-YY", New Point3d(Xinsert, Yinsert, 0), ScaraX, BTrecord, Nume_layer, colectie_goala, colectie_goala)

                    'Insereaza_Block_fara_atribute(New Point3d(Xinsert, Yinsert, 0), BTrecord, "YIN_YANG.dwg", "PIPE-YY", Nume_layer, ScaraX, ScaraY, 1, 0)
                    Trans1.Commit()
                End Using
            End If ' asta e de la rezultat vertical count



            Editor1.WriteMessage(vbLf & "Command:")



        Catch ex As Exception

            Exit Sub
            'MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("PPL_P_SYMBOL_CVR")>
    Public Sub INSERT_PIPE_SYMBOL_WITH_DEPTH_OF_COVER()

        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Dim Distanta1 As Double
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument


            Editor1 = ThisDrawing.Editor


            Dim Rezultat2 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the Nominal PIPE SIZE in Inches:")
            Rezultat2.AllowNone = True
            Dim Rezultat22 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat2)
            Dim RADIUS_IN_METERS As Double = Get_from_NPS_radius_for_pipes_from_inches_to_milimeters(Rezultat22.Value) / 1000


            Dim Point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult
            Dim PP3 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify the ground position on destination graph:")

            Dim Rezultat4 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration:")
            Rezultat4.AllowNone = True
            Rezultat4.UseDefaultValue = True
            Rezultat4.DefaultValue = 1
            Dim Rezultat44 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat4)
            Dim Exageration As Double = Rezultat44.Value
            If Exageration <= 0 Then Exageration = 1



            Dim Rezultat3 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the Depth of Cover in meters:")
            Rezultat3.AllowNone = True
            Dim Rezultat33 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat3)
            Distanta1 = Rezultat33.Value





            Dim x3, y3, Xinsert, Yinsert As Double

            PP3.AllowNone = False
            Point3 = Editor1.GetPoint(PP3)

            If Point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            x3 = Point3.Value.X
            y3 = Point3.Value.Y




            Xinsert = x3
            Yinsert = y3 - (Distanta1 * Exageration + RADIUS_IN_METERS * Exageration)

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim layer_curent As LayerTableRecord = Trans1.GetObject(ThisDrawing.Database.Clayer, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim Nume_layer As String = layer_curent.Name
                Dim layerTable As LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If layerTable.Has("TEXT") = True Then
                    Nume_layer = "TEXT"
                End If
                Dim ScaraX, ScaraY As Double
                If Exageration = 1 Then
                    ScaraX = RADIUS_IN_METERS * 2
                    ScaraY = RADIUS_IN_METERS * 2
                Else
                    ScaraY = 2 * RADIUS_IN_METERS * Exageration
                    ScaraX = ScaraY / 2
                End If

                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim colectie_goala As New Specialized.StringCollection
                InsertBlock_with_multiple_atributes("YIN_YANG.dwg", "PIPE-YY", New Point3d(Xinsert, Yinsert, 0), ScaraX, BTrecord, Nume_layer, colectie_goala, colectie_goala)
                'Insereaza_Block_fara_atribute(New Point3d(Xinsert, Yinsert, 0), BTrecord, "YIN_YANG.dwg", "PIPE-YY", Nume_layer, ScaraX, ScaraY, 1, 0)
                Trans1.Commit()
            End Using
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Exit Sub
            'MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("PPLFittingelbow", CommandFlags.UsePickSet)>
    Public Sub Show_Dual_fitting_elbow_form()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is fitting_elbow_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New fitting_elbow_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPLINS_EL_TAG")>
    Public Sub Insereaza_Elevation_Block()
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Dim Distanta1 As Double
            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument


            Editor1 = ThisDrawing.Editor
            Dim UCS_CURENT As Matrix3d = Editor1.CurrentUserCoordinateSystem



            Dim Rezultat4 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration:")
            Rezultat4.AllowNone = True
            Rezultat4.UseDefaultValue = True
            Rezultat4.DefaultValue = 1
            Dim Rezultat44 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat4)
            Dim Exageration As Double = Rezultat44.Value
            If Exageration = 0 Then Exageration = 1

            Dim Empty_array() As ObjectId

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_prompt2.MessageForAdding = vbLf & "Select a known elevation line and the label for it:"

            Object_prompt2.SingleOnly = False
            Rezultat2 = Editor1.GetSelection(Object_prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat2.Value.Count <> 2 Then
                Editor1.WriteMessage(vbLf & "Command:")
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If
123:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim x1, y1 As Double
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Specify position")
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    Exit Sub
                End If
                x1 = Point1.Value.TransformBy(UCS_CURENT).X
                y1 = Point1.Value.TransformBy(UCS_CURENT).Y

                Dim Elevatia_cunoscuta As Double = -100000
                Dim Distanta_de_la_zero As Double = -100000

                Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line

                Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj2 = Rezultat2.Value.Item(0)
                Dim Ent2 As Entity
                Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)
                Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj3 = Rezultat2.Value.Item(1)
                Dim Ent3 As Entity
                Ent3 = Obj3.ObjectId.GetObject(OpenMode.ForRead)

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent2
                    If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent3
                    If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                End If

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent2
                    If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent3
                    If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                End If

                If Elevatia_cunoscuta = -100000 Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If
                Dim x01, y01, x02, y02, dist, a As Double

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent2
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If
                    dist = ((x1 - x01) ^ 2 + (y1 - y01) ^ 2) ^ 0.5
                    a = Abs(x1 - x01)
                    Distanta_de_la_zero = (dist ^ 2 - a ^ 2) ^ 0.5
                End If


                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent3
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If
                    dist = ((x1 - x01) ^ 2 + (y1 - y01) ^ 2) ^ 0.5
                    a = Abs(x1 - x01)
                    Distanta_de_la_zero = (dist ^ 2 - a ^ 2) ^ 0.5
                End If

                If Distanta_de_la_zero = -100000 Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If


                Distanta_de_la_zero = Distanta_de_la_zero / Exageration


                Dim Elevatia1 As String

                If y1 < y01 Then
                    Elevatia1 = (Get_String_Rounded(Round(Elevatia_cunoscuta - Distanta_de_la_zero, 1), 1)).ToString
                Else
                    Elevatia1 = (Get_String_Rounded(Round(Elevatia_cunoscuta + Distanta_de_la_zero, 1), 1)).ToString
                End If



                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim layer_curent As LayerTableRecord = Trans1.GetObject(ThisDrawing.Database.Clayer, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim Nume_layer As String = layer_curent.Name
                Dim layerTable As LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If layerTable.Has("TEXT") = True Then
                    Nume_layer = "TEXT"
                End If

                Dim Colectie_atr_name As New Specialized.StringCollection
                Dim Colectie_atr_value As New Specialized.StringCollection
                Colectie_atr_name.Add("FEATURE")
                Colectie_atr_value.Add("EL.")
                Colectie_atr_name.Add("ELEVATION")
                Colectie_atr_value.Add(Elevatia1)

                InsertBlock_with_multiple_atributes("EL_TAG.dwg", "ELEVATION_TAG", New Point3d(x1, y1, 0), 1, BTrecord, Nume_layer, Colectie_atr_name, Colectie_atr_value)

                Trans1.Commit()
            End Using ' asta e de la trans1

            GoTo 123


            Editor1.WriteMessage(vbLf & "Command:")

        Catch ex As Exception
            Exit Sub
            'MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPLOW_HIGH")>
    Public Sub LOWPOINT_HIGHPOINT()
        Try
            '** added to command list
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument


            Editor1 = ThisDrawing.Editor
            Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)



            Dim Rezultat_Vline As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Prompt_Vline As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Prompt_Vline.MessageForAdding = vbLf & "Select a known Horizontal line and the label for it (ELEVATION):"

            Prompt_Vline.SingleOnly = False
            Rezultat_Vline = Editor1.GetSelection(Prompt_Vline)


            Dim Prompt_Vex As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration:")
            Prompt_Vex.DefaultValue = 1
            Prompt_Vex.AllowNone = True
            Dim Prompt_Vex4 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Prompt_Vex)
            Dim V_EXAG As Double = Prompt_Vex4.Value
            If V_EXAG = 0 Then V_EXAG = 1


            If Rezultat_Vline.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat_Vline.Value.Count <> 2 Then
                MsgBox("Your selection contains " & Rezultat_Vline.Value.Count & " objects")
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly1 As Polyline


            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim Rezultat_poly As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_prompt_poly As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_prompt_poly.MessageForAdding = vbLf & "Select the graph polyline:"

                Object_prompt_poly.SingleOnly = True
                Rezultat_poly = Editor1.GetSelection(Object_prompt_poly)
                If Not TypeOf Rezultat_poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead) Is Polyline Then
                    Editor1.SetImpliedSelection(Empty_array)
                    MsgBox("The object you selected is not a polyline")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If

                Poly1 = Rezultat_poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
            End Using


            If IsNothing(Poly1) = True Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If



            Creaza_layer("NO PLOT", 40, "", False)


123:

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                Dim layer_curent As LayerTableRecord = Trans1.GetObject(ThisDrawing.Database.Clayer, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim Nume_layer As String = layer_curent.Name
                Dim layerTable As LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If layerTable.Has("TEXT") = True Then
                    Nume_layer = "TEXT"
                End If

                Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 1st [from] point on the polyline:")
                PP_start.AllowNone = True
                Point_start = Editor1.GetPoint(PP_start)
                If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If


                Dim Point_end As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP_end As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 2nd [to] point on the polyline:")
                PP_end.AllowNone = True

                PP_end.BasePoint = Point_start.Value
                PP_end.UseBasePoint = True

                Point_end = Editor1.GetPoint(PP_end)
                If Not Point_end.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If

                Dim P1 As New Point3d
                P1 = Poly1.GetClosestPointTo(Point_start.Value, Vector3d.ZAxis, False)
                Dim P2 As New Point3d
                P2 = Poly1.GetClosestPointTo(Point_end.Value, Vector3d.ZAxis, False)

                If Poly1.GetDistAtPoint(P1) > Poly1.GetDistAtPoint(P2) Then
                    Dim p_interm As New Point3d
                    p_interm = P1
                    P1 = P2
                    P2 = p_interm
                End If
                Dim Point_y_max As New Point3d
                Dim Point_y_min As New Point3d

                If P1.Y > P2.Y Then
                    Point_y_max = P1
                    Point_y_min = P2
                Else
                    Point_y_max = P2
                    Point_y_min = P1
                End If



                Dim Index_start As Integer = Ceiling(Poly1.GetParameterAtPoint(P1))
                Dim Index_end As Integer = Floor(Poly1.GetParameterAtPoint(P2))

                If Index_start <= Index_end Then
                    For i = Index_start To Index_end
                        If Poly1.GetPointAtParameter(i).Y > Point_y_max.Y Then
                            Point_y_max = Poly1.GetPointAtParameter(i)
                        End If
                        If Poly1.GetPointAtParameter(i).Y < Point_y_min.Y Then
                            Point_y_min = Poly1.GetPointAtParameter(i)
                        End If
                    Next
                End If





                Dim Elevatia_cunoscuta As Double = -100000
                Dim Distanta_de_la_zero1 As Double = -100000

                Dim Distanta_de_la_zero2 As Double = -100000

                Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line

                Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj2 = Rezultat_Vline.Value.Item(0)
                Dim Ent2 As Entity
                Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)
                Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj3 = Rezultat_Vline.Value.Item(1)
                Dim Ent3 As Entity
                Ent3 = Obj3.ObjectId.GetObject(OpenMode.ForRead)

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent2
                    If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent3
                    If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                End If

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent2
                    If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent3
                    If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                End If

                If Elevatia_cunoscuta = -100000 Then

                    MsgBox("You have issues with elevation datum. Please check")
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If
                Dim x01, y01, x02, y02, dist1, a1 As Double

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent2
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        MsgBox("The line for elevation datum is not horizontal. Please check")
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If
                    dist1 = ((Point_y_min.X - x01) ^ 2 + (Point_y_min.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_min.X - x01)
                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5


                    dist1 = ((Point_y_max.X - x01) ^ 2 + (Point_y_max.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_max.X - x01)
                    Distanta_de_la_zero2 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                End If


                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent3
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then

                        MsgBox("The line for elevation datum is not horizontal. Please check")
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If
                    dist1 = ((Point_y_min.X - x01) ^ 2 + (Point_y_min.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_min.X - x01)
                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5


                    dist1 = ((Point_y_max.X - x01) ^ 2 + (Point_y_max.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_max.X - x01)
                    Distanta_de_la_zero2 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5
                End If

                If Distanta_de_la_zero1 = -100000 Then
                    MsgBox("You have issues with elevation datum. Please check")
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                If Distanta_de_la_zero2 = -100000 Then
                    MsgBox("You have issues with elevation datum. Please check")
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If


                Distanta_de_la_zero1 = Distanta_de_la_zero1 / V_EXAG
                Distanta_de_la_zero2 = Distanta_de_la_zero2 / V_EXAG

                Dim Elevatia1 As String

                If Point_y_min.Y < y01 Then
                    Elevatia1 = "EL: " & (Get_String_Rounded(Elevatia_cunoscuta - Distanta_de_la_zero1, 2)).ToString & "m"
                Else
                    Elevatia1 = "EL: " & (Get_String_Rounded(Elevatia_cunoscuta + Distanta_de_la_zero1, 2)).ToString & "m"
                End If

                Dim Elevatia2 As String

                If Point_y_max.Y < y01 Then
                    Elevatia2 = "EL: " & (Get_String_Rounded(Elevatia_cunoscuta - Distanta_de_la_zero2, 2)).ToString & "m"
                Else
                    Elevatia2 = "EL: " & (Get_String_Rounded(Elevatia_cunoscuta + Distanta_de_la_zero2, 2)).ToString & "m"
                End If

                Dim Mleader1 As New MLeader
                Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_y_min.TransformBy(Curent_UCS), "LOW POINT " & vbCrLf & Elevatia1, 2.5, 2.5, 2, 10, 10)
                Mleader1.Layer = "NO PLOT"

                Dim Mleader2 As New MLeader
                Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Point_y_max.TransformBy(Curent_UCS), "HIGH POINT " & vbCrLf & Elevatia2, 2.5, 2.5, 2, 10, 10)
                Mleader2.Layer = "NO PLOT"





                Dim Colectie_nume_atr As New Specialized.StringCollection
                Colectie_nume_atr.Add("FEATURE1")
                Colectie_nume_atr.Add("ELEVATION1")
                Colectie_nume_atr.Add("FEATURE2")
                Colectie_nume_atr.Add("ELEVATION2")
                Colectie_nume_atr.Add("FEATURE3")
                Colectie_nume_atr.Add("ELEVATION3")
                Colectie_nume_atr.Add("FEATURE4")
                Colectie_nume_atr.Add("ELEVATION4")

                Dim Colectie_valori_atr_low As New Specialized.StringCollection
                Colectie_valori_atr_low.Add("LOW POINT")
                Colectie_valori_atr_low.Add(Elevatia1)
                Colectie_valori_atr_low.Add("LOW POINT")
                Colectie_valori_atr_low.Add(Elevatia1)
                Colectie_valori_atr_low.Add("LOW POINT")
                Colectie_valori_atr_low.Add(Elevatia1)
                Colectie_valori_atr_low.Add("LOW POINT")
                Colectie_valori_atr_low.Add(Elevatia1)

                InsertBlock_with_multiple_atributes("HIGH-LOW.dwg", "LOW_HIGH", Point_y_min, 1, BTrecord, Nume_layer, Colectie_nume_atr, Colectie_valori_atr_low)

                Dim Colectie_valori_atr_high As New Specialized.StringCollection
                Colectie_valori_atr_high.Add("HIGH POINT")
                Colectie_valori_atr_high.Add(Elevatia2)
                Colectie_valori_atr_high.Add("HIGH POINT")
                Colectie_valori_atr_high.Add(Elevatia2)
                Colectie_valori_atr_high.Add("HIGH POINT")
                Colectie_valori_atr_high.Add(Elevatia2)
                Colectie_valori_atr_high.Add("HIGH POINT")
                Colectie_valori_atr_high.Add(Elevatia2)
                InsertBlock_with_multiple_atributes("HIGH-LOW.dwg", "LOW_HIGH", Point_y_max, 1, BTrecord, Nume_layer, Colectie_nume_atr, Colectie_valori_atr_high)




                Trans1.Commit()
            End Using ' asta e de la trans1

            GoTo 123


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception

            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("STATION_ON_GRAPH")>
    Public Sub Pick_STATION_on_graph()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
        Editor1 = ThisDrawing.Editor

        Dim Curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem

        Try
            Dim Rezultat4 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Horizontal Exaggeration:")
            Rezultat4.DefaultValue = 1
            Rezultat4.AllowNone = True
            Dim Rezultat44 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat4)
            Dim Exageration As Double = Rezultat44.Value
            If Exageration = 0 Then Exageration = 1


            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select a X position:")
            PP0.AllowNone = False
            Point0 = Editor1.GetPoint(PP0)
            If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            Dim Rezultat_chainage As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the STATION of the X point (as a number):")
            Rezultat_chainage.DefaultValue = 0
            Rezultat_chainage.AllowNone = True
            Dim Chainage_0 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_chainage)
            Dim Chainage_00 As Double = Chainage_0.Value

            Dim Rezultat_left_right As New Autodesk.AutoCAD.EditorInput.PromptKeywordOptions("")
            Rezultat_left_right.Message = vbLf & "Specify the graph Direction Left-Right or Right-Left:"
            Rezultat_left_right.Keywords.Add("LR")
            Rezultat_left_right.Keywords.Add("RL")
            Rezultat_left_right.AllowNone = False
            Rezultat_left_right.AllowArbitraryInput = False
            Rezultat_left_right.AppendKeywordsToMessage = True


            Dim Rezultat_string As Autodesk.AutoCAD.EditorInput.PromptResult = Editor1.GetKeywords(Rezultat_left_right)

            Dim Left_right As String = Rezultat_string.StringResult

            Dim FactorLR As Double = 1
            If Left_right = "RL" Then FactorLR = -1

            If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Creaza_layer("NO PLOT", 40, "", False)
                Dim Specificare_chainage As Boolean = False
                Dim Specificare_Ymin_max As Boolean = False
                Dim Ymin, Ymax As Double
                Dim Distanta_old As Double = 9999999999989.2363
                Dim Previous_default_value As Double

123:

                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select required position:")
                    PP1.AllowNone = True

                    If Specificare_chainage = False Then
                        Point1 = Editor1.GetPoint(PP1)
                    End If

                    Dim Distanta As Double
                    Dim Chainage As String
                    Dim Pt1 As New Point3d
                    Dim Pt2 As New Point3d

                    If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        Specificare_chainage = True



                        If Specificare_Ymin_max = False Then
                            Dim Rezultat__elev_lines As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                            Object_Prompt2.MessageForAdding = vbLf & "Select the top and bottom of graph:"


                            Object_Prompt2.SingleOnly = False

                            Rezultat__elev_lines = Editor1.GetSelection(Object_Prompt2)
                            If Not Rezultat__elev_lines.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                            Dim Assigned As Boolean = False

                            For i = 1 To Rezultat__elev_lines.Value.Count
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat__elev_lines.Value.Item(i - 1)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                                    Dim Linie1 As Line = Ent1
                                    If Abs(Linie1.StartPoint.Y - Linie1.EndPoint.Y) < 0.02 Then
                                        If i = 1 Then
                                            Assigned = True
                                            Ymin = Linie1.StartPoint.Y
                                            Ymax = Linie1.StartPoint.Y
                                        Else
                                            If Linie1.StartPoint.Y > Ymax Then Ymax = Linie1.StartPoint.Y
                                            If Linie1.StartPoint.Y < Ymin Then Ymin = Linie1.StartPoint.Y
                                        End If
                                    End If
                                End If
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim PLinie1 As Polyline = Ent1
                                    If PLinie1.StartPoint.Y = PLinie1.EndPoint.Y And PLinie1.NumberOfVertices = 2 Then
                                        If i = 1 And Assigned = False Then
                                            Ymin = PLinie1.StartPoint.Y
                                            Ymax = PLinie1.StartPoint.Y
                                        Else
                                            If PLinie1.StartPoint.Y > Ymax Then Ymax = PLinie1.StartPoint.Y
                                            If PLinie1.StartPoint.Y < Ymin Then Ymin = PLinie1.StartPoint.Y
                                        End If
                                    End If
                                End If
                            Next
                            Specificare_Ymin_max = True
                        End If

                        Dim Rezultat_chainage1 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the desired STATION:")
                        'Rezultat_chainage1.DefaultValue = Previous_default_value
                        Rezultat_chainage1.AllowNone = True

                        Dim Chainage_1 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_chainage1)

                        If Not Chainage_1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Dim Chainage_11 As Double = Chainage_1.Value
                        Previous_default_value = Chainage_11
                        Distanta = Chainage_11
                        If Distanta = Distanta_old Then
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Distanta_old = Distanta



                        Pt1 = New Point3d(Point0.Value.X + FactorLR * (Distanta - Chainage_00) * Exageration, Ymin, 0)
                        Pt2 = New Point3d(Point0.Value.X + FactorLR * (Distanta - Chainage_00) * Exageration, Ymax, 0)



                    Else

                        Pt1 = Point1.Value
                        Pt2 = New Point3d(Point1.Value.X, Point0.Value.Y, 0)

                        Distanta = Chainage_00 + FactorLR * (Point1.Value.X - Point0.Value.X) / Exageration


                    End If

                    Chainage = Get_chainage_feet_from_double(Distanta, 1)

                    If Chainage = "-0+000.0" Then Chainage = "0+000.0"

                    Dim Mtext_chainage_25 As New MText
                    Mtext_chainage_25.Contents = Chainage
                    Mtext_chainage_25.Location = Pt2
                    Mtext_chainage_25.Attachment = AttachmentPoint.BottomLeft
                    Mtext_chainage_25.TextHeight = 2.5
                    Mtext_chainage_25.Rotation = PI / 2
                    Mtext_chainage_25.ColorIndex = 256
                    BTrecord.AppendEntity(Mtext_chainage_25)
                    Trans1.AddNewlyCreatedDBObject(Mtext_chainage_25, True)

                    Creaza_Mleader_nou_fara_UCS_transform(Pt2.TransformBy(Curent_ucs_matrix), Chainage, 2.5, 1, 2, 10, 10)

                    Dim Mtext_chainage_REF_survey As New MText
                    Mtext_chainage_REF_survey.Contents = Chainage
                    Mtext_chainage_REF_survey.Location = Pt1
                    Mtext_chainage_REF_survey.Attachment = AttachmentPoint.MiddleLeft
                    Mtext_chainage_REF_survey.TextHeight = 1
                    Mtext_chainage_REF_survey.Rotation = PI / 2
                    Mtext_chainage_REF_survey.Layer = "NO PLOT"
                    Mtext_chainage_REF_survey.ColorIndex = 256
                    BTrecord.AppendEntity(Mtext_chainage_REF_survey)
                    Trans1.AddNewlyCreatedDBObject(Mtext_chainage_REF_survey, True)


                    Dim Line1 As New Line(Pt1, Pt2)
                    Line1.Layer = "NO PLOT"
                    BTrecord.AppendEntity(Line1)
                    Trans1.AddNewlyCreatedDBObject(Line1, True)
                    Trans1.Commit()

                    GoTo 123
                End Using

            Else
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub

            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub



    <CommandMethod("PPL_PT_CH")>
    Public Sub point_at_chainage()

        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If
            Dim Poly1 As Polyline

            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly1 = Ent1

                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_zero As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 0+000:" & vbCrLf & "Or press enter for start of the polyline")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)
                            If Not Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_zero = Poly1.GetClosestPointTo(Poly1.StartPoint, Vector3d.ZAxis, False)
                            Else
                                'aici am tratat ucs-ul 
                                Point_zero = Poly1.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            End If

                            Dist_from_start_for_zero = Poly1.GetDistAtPoint(Point_zero)

                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_zero As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 0+000:" & vbCrLf & "Or press enter for start of the polyline")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)
                            If Not Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_zero = Poly3D.GetClosestPointTo(Poly3D.StartPoint, Vector3d.ZAxis, False)
                            Else
                                'aici am tratat ucs-ul 
                                Point_zero = Poly3D.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            End If

                            Dist_from_start_for_zero = Poly3D.GetDistAtPoint(Point_zero)

                            Trans1.Commit()
                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim String1 As Autodesk.AutoCAD.EditorInput.PromptStringOptions
                String1 = New Autodesk.AutoCAD.EditorInput.PromptStringOptions(vbLf & "Specify station:")
                String1.AllowSpaces = True

                Dim Descriptia As Autodesk.AutoCAD.EditorInput.PromptResult = Editor1.GetString(String1)

                If Descriptia.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    Exit Sub
                End If

                Dim Ch_result As String = Descriptia.StringResult
                Ch_result = Replace(Ch_result, "+", "")
                Ch_result = Replace(Ch_result, " ", "")
                If IsNumeric(Ch_result) = False Then
                    MsgBox("Chainage is not specified correctly")
                    Exit Sub

                End If
                Dim Chainage As Double = CDbl(Ch_result)

                If Dist_from_start_for_zero + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If
                If IsNothing(Poly1) = False Then
                    Point_on_poly = Poly1.GetPointAtDist(Dist_from_start_for_zero + Chainage)
                End If
                If IsNothing(Poly3D) = False Then
                    Point_on_poly = Poly3D.GetPointAtDist(Dist_from_start_for_zero + Chainage)
                End If

                Dim Chainage_string As String = Get_chainage_from_double(Chainage, 1)
                If Chainage_string = "-0+000.0" Then Chainage_string = "0+000.0"

                If IsNothing(Point_on_poly) = False Then
                    Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 5, 2.5, 5, 10, 10)
                End If

                Trans1.Commit()

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("PPL_CH_PT")>
    Public Sub Creaza_CHAINAGE_LABEL_ON_THE_2DPOLYLINE()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly1 As Polyline
            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly1 = Ent1

                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_zero As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 0+000:" & vbCrLf & "Or press enter for start of the polyline")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)
                            If Not Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_zero = Poly1.StartPoint
                            Else
                                'aici am tratat ucs-ul 
                                Point_zero = Poly1.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            End If

                            Dist_from_start_for_zero = Poly1.GetDistAtPoint(Point_zero)

                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_zero As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 0+000:" & vbCrLf & "Or press enter for start of the polyline")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)
                            If Not Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_zero = Poly3D.StartPoint
                            Else
                                'aici am tratat ucs-ul 
                                Point_zero = Poly3D.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            End If

                            Dist_from_start_for_zero = Poly3D.GetDistAtPoint(Point_zero)

                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the same polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly1) = False Then
                    Point_on_poly = Poly1.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly1.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D) = False Then
                    Point_on_poly = Poly3D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly3D.GetDistAtPoint(Point_on_poly)
                End If


                Dim Chainage As Double = Distanta_pana_la_xing - Dist_from_start_for_zero




                If Dist_from_start_for_zero + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired chainage point is not matching.")
                    Exit Sub
                End If




                Dim Chainage_string As String = Get_chainage_from_double(Chainage, 1)
                If Chainage_string = "-0+000.0" Then Chainage_string = "0+000.0"

                Dim Mleader1 As New MLeader

                If IsNothing(Point_on_poly) = False Then
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 5, 2.5, 5, 10, 10)
                End If


                Trans1.Commit()

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("PPLHDD")>
    Public Sub Show_hdd_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is HDD_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New HDD_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    <CommandMethod("PPL_UBOLT3D", CommandFlags.UsePickSet)>
    Public Sub Show_uBOLT_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is _3d_Ubolt_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New _3d_Ubolt_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("PPL_CLAMP3D", CommandFlags.UsePickSet)>
    Public Sub _3d_clamp()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is _3d_clamp_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New _3d_clamp_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPL_TRANSCANADA")>
    Public Sub Show_TRANSCANADA_LAYERS_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Transcanada_layers_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Transcanada_layers_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("PPLDIM")>
    Public Sub Show_dim_change_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Dimension_change_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Dimension_change_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_CH2XL")>
    Public Sub Creaza_CHAINAGE_LABEL_ON_dual_2DPOLY_then_transfer_to_EXCEL()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select Old polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select New polyline:"

            Object_Prompt2.SingleOnly = True

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Poly1 As Polyline
            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double = 0

            Dim Poly2 As Polyline
            Dim Point_on_poly2 As New Point3d

            Dim Dist_from_start_for_zero2 As Double = 0

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat2.Value.Item(0)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Poly1 = Ent1
                            Poly2 = Ent2
                            Trans1.Commit()
                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = get_new_worksheet_from_Excel()
            W1.Range("A1").Value = "OLD"
            W1.Range("B1").Value = "NEW"

            Dim Index_XL As Double = 2
            Dim Celula_A As Microsoft.Office.Interop.Excel.Range
            Dim Celula_B As Microsoft.Office.Interop.Excel.Range

1234:
            Celula_A = W1.Range("A" & Index_XL)
            Celula_B = W1.Range("B" & Index_XL)
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please Pick a point on the OLD polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please Pick a point on the NEW polyline:")
                Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP2.AllowNone = False
                Point2 = Editor1.GetPoint(PP2)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If


                Point_on_poly = Poly1.GetClosestPointTo(Point1.Value.TransformBy(CurentUCSmatrix), Poly1.Normal, False)
                Dim Distanta_pana_la_xing As Double = Poly1.GetDistAtPoint(Point_on_poly)
                Dim Chainage As Double = Distanta_pana_la_xing - Dist_from_start_for_zero
                If Dist_from_start_for_zero + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired chainage point is not matching.")
                    Exit Sub
                End If
                Dim Chainage_string As String = Get_chainage_from_double(Chainage, 1)
                If Chainage_string = "-0+000.0" Then Chainage_string = "0+000.0"
                Dim Mleader1 As New MLeader
                Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 5, 2.5, 5, 10, 10)

                Point_on_poly2 = Poly2.GetClosestPointTo(Point2.Value.TransformBy(CurentUCSmatrix), Poly1.Normal, False)
                Dim Distanta_pana_la_xing2 As Double = Poly2.GetDistAtPoint(Point_on_poly2)
                Dim Chainage2 As Double = Distanta_pana_la_xing2 - Dist_from_start_for_zero2
                If Dist_from_start_for_zero2 + Chainage2 < 0 Then
                    MsgBox("The 0+000 position and your desired chainage point is not matching.")
                    Exit Sub
                End If
                Dim Chainage_string2 As String = Get_chainage_from_double(Chainage2, 1)
                If Chainage_string2 = "-0+000.0" Then Chainage_string2 = "0+000.0"
                Dim Mleader2 As New MLeader
                Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly2, Chainage_string2, 5, 2.5, 5, 10, 10)
                Mleader2.ColorIndex = 1



                Trans1.Commit()
                Celula_A.Value = Chainage_string
                Celula_B.Value = Chainage_string2
                Index_XL = Index_XL + 1
                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub


    Public Sub Show_find_replace_from_excel_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is FIND_REPLACE_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New FIND_REPLACE_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    <CommandMethod("PPL_W2XL")>
    Public Sub Show_READ_TEXTW_2_excel_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Read_from_acad_w_excel_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Read_from_acad_w_excel_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPL_PT_INS")>
    Public Sub Show_POINT_INSERTOR_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Point_insertor_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Point_insertor_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPL_CH_PT_RRT")>
    Public Sub Creaza_CHAINAGE_LABEL_ON_THE_2DPOLYLINE_for_REROUTE()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select OLD polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select NEW polyline:"

            Object_Prompt2.SingleOnly = True

            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Poly1 As Polyline
            Dim Poly3D As Polyline3d

            Dim Poly2 As Polyline
            Dim Poly3D2 As Polyline3d


            Dim Point_on_poly As New Point3d

            Dim Chainage_at_common_point_old As Double
            Dim Chainage_at_common_point_new As Double
            Dim Diferenta_chainage As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat2.Value.Item(0)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly1 = Ent1
                            Poly2 = Ent2

                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_zero_old As New Point3d
                            Dim Point_zero_new As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select reroute start point:")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)
                            If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_zero_old = Poly1.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            Else
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                            Chainage_at_common_point_old = Poly1.GetDistAtPoint(Point_zero_old)
                            Point_zero_new = Poly2.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            Chainage_at_common_point_new = Poly2.GetDistAtPoint(Point_zero_new)
                            Diferenta_chainage = Chainage_at_common_point_new - Chainage_at_common_point_old

                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d And TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1
                            Poly3D2 = Ent2

                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_zero_old As New Point3d
                            Dim Point_zero_new As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select reroute start point:")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)


                            If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_zero_old = Poly3D.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            Else
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                            Chainage_at_common_point_old = Poly3D.GetDistAtPoint(Point_zero_old)
                            Point_zero_new = Poly3D2.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            Chainage_at_common_point_new = Poly3D2.GetDistAtPoint(Point_zero_new)
                            Diferenta_chainage = Chainage_at_common_point_new - Chainage_at_common_point_old

                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polylines")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the reroute:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly2) = False Then
                    Point_on_poly = Poly2.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly2.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D2) = False Then
                    Point_on_poly = Poly3D2.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly3D2.GetDistAtPoint(Point_on_poly)
                End If


                Dim Chainage As Double = Distanta_pana_la_xing - Diferenta_chainage




                Dim Chainage_string As String = Get_chainage_from_double(Chainage, 1)
                If Chainage_string = "-0+000.0" Then Chainage_string = "0+000.0"

                Dim Mleader1 As New MLeader

                If IsNothing(Point_on_poly) = False Then
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 0.5, 0.1, 0.5, 11, 5)
                End If


                Trans1.Commit()

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("AV2P")>
    Public Sub Align_view_and_UCS_to_points()
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem


        Try

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                Dim Point_1 As New Point3d
                Dim Point_2 As New Point3d


                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select first point:")
                PP1.AllowNone = True
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If
                Point_1 = Point1.Value

                Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select second point:")
                PP2.AllowNone = True
                PP2.UseBasePoint = True
                PP2.BasePoint = Point1.Value
                Point2 = Editor1.GetPoint(PP2)
                If Not Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If
                Point_2 = Point2.Value


                Point_1 = Point_1.TransformBy(CurentUCSmatrix)
                Point_2 = Point_2.TransformBy(CurentUCSmatrix)

                Dim kd As Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor = New Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor()
                kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"))
                Dim Cvport As Integer = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"))
                Using GraphicsManager As Autodesk.AutoCAD.GraphicsSystem.Manager = ThisDrawing.GraphicsManager
                    Using ViewC As Autodesk.AutoCAD.GraphicsSystem.View = GraphicsManager.ObtainAcGsView(Cvport, kd)
                        'view.ZoomExtents(Ent1.GeometricExtents.MaxPoint, Ent1.GeometricExtents.MinPoint)

                        'GraphicsManager.SetViewportFromView(Cvport, view, True, True, False)
                    End Using
                End Using


                Dim Ucs_table As UcsTable = Trans1.GetObject(ThisDrawing.Database.UcsTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim View_Table As ViewTable = Trans1.GetObject(ThisDrawing.Database.ViewTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                Dim Ratio As Double = 1.3 'ThisDrawing.Editor.GetCurrentView.Width / ThisDrawing.Editor.GetCurrentView.Height
                Dim Latime_view As Double = 1.3 * Point_1.GetVectorTo(Point_2).Length
                Dim Ucs1 As UcsTableRecord
                Dim ucs_NAME As String = "1123"

                If Ucs_table.Has(ucs_NAME) = False Then
                    Ucs1 = New UcsTableRecord
                    Ucs1.Name = ucs_NAME
                    Ucs_table.UpgradeOpen()
                    Ucs_table.Add(Ucs1)
                    Trans1.AddNewlyCreatedDBObject(Ucs1, True)
                    Ucs1.Origin = Point_1
                    Ucs1.XAxis = Point_1.GetVectorTo(Point_2)
                    Ucs1.YAxis = Vector3d.ZAxis.CrossProduct(Point_1.GetVectorTo(Point_2))

                    Dim ViewportTableRecord1 As ViewportTableRecord
                    ViewportTableRecord1 = Trans1.GetObject(ThisDrawing.Editor.ActiveViewportId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    ViewportTableRecord1.IconAtOrigin = True
                    ViewportTableRecord1.IconEnabled = True
                    ViewportTableRecord1.SetUcs(Ucs1.ObjectId)
                    ThisDrawing.Editor.UpdateTiledViewportsFromDatabase()
                Else
                    Ucs1 = Ucs_table(ucs_NAME).GetObject(OpenMode.ForWrite)
                    Ucs1.Origin = Point_1
                    Ucs1.XAxis = Point_1.GetVectorTo(Point_2)
                    Ucs1.YAxis = Vector3d.ZAxis.CrossProduct(Point_1.GetVectorTo(Point_2))

                    Dim ViewportTableRecord1 As ViewportTableRecord
                    ViewportTableRecord1 = Trans1.GetObject(ThisDrawing.Editor.ActiveViewportId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    ViewportTableRecord1.IconAtOrigin = True
                    ViewportTableRecord1.IconEnabled = True
                    ViewportTableRecord1.SetUcs(Ucs1.ObjectId)
                    ThisDrawing.Editor.UpdateTiledViewportsFromDatabase()

                End If

                Dim PointM As New Point3d((Point_1.X + Point_2.X) / 2, (Point_1.Y + Point_2.Y) / 2, 0)


                Dim Rotatie As Double = Vector3d.YAxis.GetAngleTo(Ucs1.YAxis)
                Dim Vector1 As Vector3d = Vector3d.YAxis
                Dim Vector2 As Vector3d = Ucs1.YAxis

                If Vector1.AngleOnPlane(New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)) < Vector2.AngleOnPlane(New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)) Then
                    Rotatie = 2 * PI - Rotatie
                End If

                Dim view1 As ViewTableRecord

                If View_Table.Has(ucs_NAME) = False Then
                    View_Table.UpgradeOpen()
                    view1 = New ViewTableRecord
                    view1.CenterPoint = New Point2d(0, 0)
                    view1.Target = PointM
                    view1.ViewTwist = Rotatie
                    view1.Width = Latime_view
                    view1.Height = Latime_view / Ratio
                    view1.Name = ucs_NAME
                    View_Table.Add(view1)
                    Trans1.AddNewlyCreatedDBObject(view1, True)
                Else
                    view1 = View_Table(ucs_NAME).GetObject(OpenMode.ForWrite)
                    view1.CenterPoint = New Point2d(0, 0)
                    view1.Target = PointM
                    view1.ViewTwist = Rotatie
                    view1.Width = Latime_view
                    view1.Height = Latime_view / Ratio
                End If


                ThisDrawing.Editor.SetCurrentView(view1)


                Trans1.Commit()
            End Using



            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_RECHAINAGE")>
    Public Sub Show_POINT_RECHAINAGE_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Rechainage_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Rechainage_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub



    <CommandMethod("PPL_MATERIALS")>
    Public Sub Show_materials_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Material_calc_for_alignment_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Material_calc_for_alignment_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_chaincalc")>
    Public Sub Show_chainage_calc_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Chainage_operations_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Chainage_operations_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("ATR_TAG2TEXT", CommandFlags.UsePickSet)>
    Public Sub ATR_TAG2TEXT()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)




                        For i = 0 To Rezultat1.Value.Count - 1

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is AttributeDefinition Then




                                Dim Attrib1 As AttributeDefinition = Ent1
                                Dim Ent2 As New DBText
                                Ent2.TextString = Attrib1.Tag
                                Ent2.Layer = Attrib1.Layer
                                Ent2.Rotation = Attrib1.Rotation
                                Ent2.Height = Attrib1.Height
                                'Ent2.TextStyleId = Attrib1.TextStyleId
                                Ent2.ColorIndex = Attrib1.ColorIndex
                                Ent2.LineWeight = Attrib1.LineWeight
                                Ent2.Position = Attrib1.Position
                                BTrecord.AppendEntity(Ent2)
                                Trans1.AddNewlyCreatedDBObject(Ent2, True)
                                Ent1.UpgradeOpen()
                                Ent1.Erase()

                            End If










                        Next


                        Trans1.Commit()
                        Editor1.Regen()


                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try


    End Sub


    <CommandMethod("atsync", CommandFlags.UsePickSet)>
    Public Sub Sync_attrib_from_block()
        Try
            Dim Nume_block As New Specialized.StringCollection
            Using lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


                Rezultat1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.SelectImplied

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                Else
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select the Block references:"
                    Object_Prompt.SingleOnly = False
                    Rezultat1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetSelection(Object_Prompt)
                End If






                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction
                        For i = 0 To Rezultat1.Value.Count - 1
                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                Dim Block1 As BlockReference = Ent1
                                If Block1.AttributeCollection.Count > 0 Then
                                    Nume_block.Add(Block1.Name)
                                End If

                            End If
                        Next

                        Trans1.Commit()
                    End Using 'asta e de la trans1

                End If
                If Nume_block.Count > 0 Then
                    Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    For i = 0 To Nume_block.Count - 1
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            Dim BlocktableRec1 As BlockTableRecord = BlockTable_data1.Item(Nume_block(i)).GetObject(OpenMode.ForWrite)
                            SynchronizeAttributes_db_diferit(BlocktableRec1, Trans1)
                            Trans1.Commit()
                        End Using
                    Next

                End If
            End Using

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            MsgBox(ex.Message)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
        End Try
    End Sub

    <CommandMethod("PPL_surveyband")>
    Public Sub Show_DEFLECTIONS_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Survey_band_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Survey_band_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_CSFCHAINAGE")>
    Public Sub CALCULEAZA_CHAINAGE_WITH_CSF()
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select 3D polyline:"
            Object_Prompt2.SingleOnly = True
            Rezultat2 = Editor1.GetSelection(Object_Prompt2)
            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly3d As Polyline3d
            Dim Point_on_poly As New Point3d
            Dim Poly2D As New Polyline

            If Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                If IsNothing(Rezultat2) = False Then
                    Using lock1 As DocumentLock = ThisDrawing.LockDocument
                        Dim Data_table1 As New System.Data.DataTable
                        Data_table1.Columns.Add("TEXT325", GetType(DBText))
                        Dim Index1 As Double = 0
                        Dim Data_table2 As New System.Data.DataTable
                        Data_table2.Columns.Add("TEXT0", GetType(DBText))
                        Dim Index2 As Double = 0
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj2 = Rezultat2.Value.Item(0)
                            Dim Ent2 As Entity
                            Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                                Poly3d = Ent2
                            Else
                                Editor1.WriteMessage("No 3d Polyline")
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If
                            For Each ObjID In BTrecord
                                Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If TypeOf DBobject Is DBText Then
                                    Dim Text1 As DBText = DBobject
                                    If Text1.Layer = Poly3d.Layer Then
                                        If Text1.Rotation > 3 * PI / 2 Then
                                            Data_table1.Rows.Add()
                                            Data_table1.Rows(Index1).Item("TEXT325") = Text1
                                            Index1 = Index1 + 1
                                        End If
                                        If Text1.Rotation >= 0 And Text1.Rotation < PI / 4 Then
                                            Data_table2.Rows.Add()
                                            Data_table2.Rows(Index2).Item("TEXT0") = Text1
                                            Index2 = Index2 + 1
                                        End If

                                    End If
                                End If
                            Next


                            Dim Index2d As Double = 0
                            For Each ObjId As ObjectId In Poly3d
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)

                                Index2d = Index2d + 1
                            Next
                            Poly2D.Elevation = 0
                        End Using


123:

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select Point:")
                            PP1.AllowNone = True
                            Point1 = Editor1.GetPoint(PP1)
                            If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                Trans1.Commit()
                                Exit Sub
                            End If




                            Dim Point_on_poly2D As New Point3d
                            Point_on_poly2D = Poly2D.GetClosestPointTo(New Point3d(Point1.Value.X, Point1.Value.Y, 0).TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                            Dim Param1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2D)
                            Point_on_poly = Poly3d.GetPointAtParameter(Param1)

                            Dim New_ch As Double = Get_chainage_with_CSF(Poly3d, Point_on_poly, Data_table2, Data_table1)
                            Dim New_chainage As String = "X = " & Round(Point_on_poly.X, 2) & vbCrLf &
                                                        "Y = " & Round(Point_on_poly.Y, 2) & vbCrLf &
                                                        "Z = " & Round(Point_on_poly.Z, 2) & vbCrLf &
                                                        Get_chainage_from_double(New_ch, 1)


                            Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, New_chainage, 0.5, 0.5, 0.5, 11, 3.5)

                            Trans1.Commit()

                        End Using
                        GoTo 123
                    End Using
                End If
            End If

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPL_WT2XL", CommandFlags.UsePickSet)>
    Public Sub Write_column_to_excel()
        If isSECURE() = False Then Exit Sub

        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Line_prompt0 As New Autodesk.AutoCAD.EditorInput.PromptStringOptions("Specify column in Excel")
            Dim Rezultat0 As Autodesk.AutoCAD.EditorInput.PromptResult = Editor1.GetString(Line_prompt0)

            If Rezultat0.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If

            Dim Line_prompt00 As New Autodesk.AutoCAD.EditorInput.PromptIntegerOptions("Specify start row in Excel")
            Line_prompt00.AllowNegative = False
            Line_prompt00.AllowZero = False
            Line_prompt00.AllowNone = True
            Dim Rezultat00 As Autodesk.AutoCAD.EditorInput.PromptIntegerResult = Editor1.GetInteger(Line_prompt00)

            If Rezultat00.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select 1 column of mtext objects"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
                If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Exit Sub
                End If
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                If IsNothing(Rezultat1) = False Then
                    Dim Area1 As Double = 0
                    Using Trans2 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim Selset1 As Autodesk.AutoCAD.EditorInput.SelectionSet = Rezultat1.Value
                        Dim Ent1 As Autodesk.AutoCAD.DatabaseServices.Entity
                        Dim Colectie_string As New Specialized.StringCollection
                        Dim Colectie_pozitie_Y As New DoubleCollection

                        For i = 0 To Selset1.Count - 1
                            Ent1 = Selset1(i).ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                Using Mtext1 As MText = Ent1
                                    Colectie_string.Add(Mtext1.Text)
                                    Colectie_pozitie_Y.Add(Round(Mtext1.Location.Y, 2))
                                    'MsgBox(Mtext1.Location.Y)
                                End Using
                            End If

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Using text1 As DBText = Ent1
                                    Colectie_string.Add(text1.TextString)
                                    Colectie_pozitie_Y.Add(Round(text1.Position.Y, 2))
                                    'MsgBox(Mtext1.Location.Y)
                                End Using
                            End If
                        Next

                        Dim Y_Array(Colectie_pozitie_Y.Count - 1) As Double
                        Dim String_Array(Colectie_string.Count - 1) As String

                        For i = 0 To Colectie_pozitie_Y.Count - 1
                            Y_Array(i) = Colectie_pozitie_Y(i)
                            String_Array(i) = Colectie_string(i)
                        Next
                        Array.Sort(Y_Array, String_Array)
                        Array.Reverse(String_Array)

                        Dim w1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
                        Dim Start_row As Integer = Rezultat00.Value
                        For i = 0 To Colectie_string.Count - 1
                            Dim Celula1 As Microsoft.Office.Interop.Excel.Range = w1.Range(Rezultat0.StringResult.ToUpper & (i + Start_row))
                            Celula1.Value = String_Array(i)
                        Next
                    End Using
                End If
            End If

123:


            Dim Empty_array1() As ObjectId
            Editor1.SetImpliedSelection(Empty_array1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPLOW_HIGH_ALG")>
    Public Sub LOWPOINT_HIGHPOINT_ALIGNMENT()
        Try
            '** added to command list
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document

            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor

            ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Editor1 = ThisDrawing.Editor
            Dim Curent_UCS As Matrix3d = Editor1.CurrentUserCoordinateSystem

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Dim Rezultat4 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Horizontal Exaggeration:")
            Rezultat4.DefaultValue = 1
            Rezultat4.AllowNone = True
            Dim Rezultat44 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat4)
            Dim HExageration As Double = Rezultat44.Value
            If HExageration = 0 Then HExageration = 1


            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult

            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select a X position:")
            PP0.AllowNone = False
            Point0 = Editor1.GetPoint(PP0)
            If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            Dim Rezultat_chainage As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the chainage of the X point (as a number):")
            Rezultat_chainage.DefaultValue = 0
            Rezultat_chainage.AllowNone = True
            Dim Chainage_0 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat_chainage)
            Dim Chainage_00 As Double = 0
            If Chainage_0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Chainage_00 = Chainage_0.Value
            End If



            Dim Rezultat_left_right As New Autodesk.AutoCAD.EditorInput.PromptKeywordOptions("")
            Rezultat_left_right.Message = vbLf & "Specify the graph Direction Left-Right or Right-Left:"
            Rezultat_left_right.Keywords.Add("LR")
            Rezultat_left_right.Keywords.Add("RL")
            Rezultat_left_right.AllowNone = False
            Rezultat_left_right.AllowArbitraryInput = False
            Rezultat_left_right.AppendKeywordsToMessage = True


            Dim Rezultat_string As Autodesk.AutoCAD.EditorInput.PromptResult = Editor1.GetKeywords(Rezultat_left_right)

            Dim Left_right As String = Rezultat_string.StringResult

            Dim FactorLR As Double = 1
            If Left_right = "RL" Then FactorLR = -1


            Dim Rezultat_Vline As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Prompt_Vline As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Prompt_Vline.MessageForAdding = vbLf & "Select a known Horizontal line and the label for it (ELEVATION):"

            Prompt_Vline.SingleOnly = False
            Rezultat_Vline = Editor1.GetSelection(Prompt_Vline)


            Dim Prompt_Vex As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Vertical Exaggeration:")
            Prompt_Vex.DefaultValue = 1
            Prompt_Vex.AllowNone = True
            Dim Prompt_Vex4 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Prompt_Vex)
            Dim V_EXAG As Double = Prompt_Vex4.Value
            If V_EXAG = 0 Then V_EXAG = 1


            If Rezultat_Vline.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat_Vline.Value.Count <> 2 Then
                MsgBox("Your selection contains " & Rezultat_Vline.Value.Count & " objects")
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly1 As Polyline


            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim Rezultat_poly As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_prompt_poly As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_prompt_poly.MessageForAdding = vbLf & "Select the graph polyline:"

                Object_prompt_poly.SingleOnly = True
                Rezultat_poly = Editor1.GetSelection(Object_prompt_poly)
                If Not TypeOf Rezultat_poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead) Is Polyline Then
                    Editor1.SetImpliedSelection(Empty_array)
                    MsgBox("The object you selected is not a polyline")
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If

                Poly1 = Rezultat_poly.Value.Item(0).ObjectId.GetObject(OpenMode.ForRead)
            End Using


            If IsNothing(Poly1) = True Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If



            Creaza_layer("NO PLOT", 40, "", False)
            Dim nR_TEST_SECTION = 1

123:

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                Dim layer_curent As LayerTableRecord = Trans1.GetObject(ThisDrawing.Database.Clayer, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim Nume_layer As String = layer_curent.Name
                Dim layerTable As LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If layerTable.Has("TEXT") = True Then
                    Nume_layer = "TEXT"
                End If

                Dim Point_start As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP_start As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 1st [from] point on the polyline:")
                PP_start.AllowNone = True
                Point_start = Editor1.GetPoint(PP_start)
                If Not Point_start.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If


                Dim Point_end As Autodesk.AutoCAD.EditorInput.PromptPointResult

                Dim PP_end As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 2nd [to] point on the polyline:")
                PP_end.AllowNone = True

                PP_end.BasePoint = Point_start.Value
                PP_end.UseBasePoint = True

                Point_end = Editor1.GetPoint(PP_end)
                If Not Point_end.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If

                Dim P1 As New Point3d
                P1 = Poly1.GetClosestPointTo(Point_start.Value, Vector3d.ZAxis, False)
                Dim P2 As New Point3d
                P2 = Poly1.GetClosestPointTo(Point_end.Value, Vector3d.ZAxis, False)

                If Poly1.GetDistAtPoint(P1) > Poly1.GetDistAtPoint(P2) Then
                    Dim p_interm As New Point3d
                    p_interm = P1
                    P1 = P2
                    P2 = p_interm
                End If
                Dim Point_y_max As New Point3d
                Dim Point_y_min As New Point3d

                If P1.Y > P2.Y Then
                    Point_y_max = P1
                    Point_y_min = P2
                Else
                    Point_y_max = P2
                    Point_y_min = P1
                End If



                Dim Index_start As Integer = Ceiling(Poly1.GetParameterAtPoint(P1))
                Dim Index_end As Integer = Floor(Poly1.GetParameterAtPoint(P2))

                If Index_start <= Index_end Then
                    For i = Index_start To Index_end
                        If Poly1.GetPointAtParameter(i).Y > Point_y_max.Y Then
                            Point_y_max = Poly1.GetPointAtParameter(i)
                        End If
                        If Poly1.GetPointAtParameter(i).Y < Point_y_min.Y Then
                            Point_y_min = Poly1.GetPointAtParameter(i)
                        End If
                    Next
                End If





                Dim Elevatia_cunoscuta As Double = -100000
                Dim Distanta_de_la_zero1 As Double = -100000

                Dim Distanta_de_la_zero2 As Double = -100000

                Dim mText_cunoscut As Autodesk.AutoCAD.DatabaseServices.MText
                Dim Text_cunoscut As Autodesk.AutoCAD.DatabaseServices.DBText
                Dim Linia_cunoscuta As Autodesk.AutoCAD.DatabaseServices.Line

                Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj2 = Rezultat_Vline.Value.Item(0)
                Dim Ent2 As Entity
                Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)
                Dim Obj3 As Autodesk.AutoCAD.EditorInput.SelectedObject
                Obj3 = Rezultat_Vline.Value.Item(1)
                Dim Ent3 As Entity
                Ent3 = Obj3.ObjectId.GetObject(OpenMode.ForRead)

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent2
                    If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                    mText_cunoscut = Ent3
                    If IsNumeric(mText_cunoscut.Text) = True Then Elevatia_cunoscuta = CDbl(mText_cunoscut.Text)
                End If

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent2
                    If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                End If

                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                    Text_cunoscut = Ent3
                    If IsNumeric(Text_cunoscut.TextString) = True Then Elevatia_cunoscuta = CDbl(Text_cunoscut.TextString)
                End If

                If Elevatia_cunoscuta = -100000 Then

                    MsgBox("You have issues with elevation datum. Please check")
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If
                Dim x01, y01, x02, y02, dist1, a1 As Double

                If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent2
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then
                        MsgBox("The line for elevation datum is not horizontal. Please check")
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If
                    dist1 = ((Point_y_min.X - x01) ^ 2 + (Point_y_min.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_min.X - x01)
                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5


                    dist1 = ((Point_y_max.X - x01) ^ 2 + (Point_y_max.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_max.X - x01)
                    Distanta_de_la_zero2 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5

                End If


                If TypeOf Ent3 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                    Linia_cunoscuta = Ent3
                    x01 = Linia_cunoscuta.StartPoint.X
                    y01 = Linia_cunoscuta.StartPoint.Y
                    x02 = Linia_cunoscuta.EndPoint.X
                    y02 = Linia_cunoscuta.EndPoint.Y
                    If Abs(y01 - y02) > 0.001 Then

                        MsgBox("The line for elevation datum is not horizontal. Please check")
                        Editor1.SetImpliedSelection(Empty_array)
                        Exit Sub
                    End If
                    dist1 = ((Point_y_min.X - x01) ^ 2 + (Point_y_min.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_min.X - x01)
                    Distanta_de_la_zero1 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5


                    dist1 = ((Point_y_max.X - x01) ^ 2 + (Point_y_max.Y - y01) ^ 2) ^ 0.5
                    a1 = Abs(Point_y_max.X - x01)
                    Distanta_de_la_zero2 = (dist1 ^ 2 - a1 ^ 2) ^ 0.5
                End If

                If Distanta_de_la_zero1 = -100000 Then
                    MsgBox("You have issues with elevation datum. Please check")
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If

                If Distanta_de_la_zero2 = -100000 Then
                    MsgBox("You have issues with elevation datum. Please check")
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If


                Distanta_de_la_zero1 = Distanta_de_la_zero1 / V_EXAG
                Distanta_de_la_zero2 = Distanta_de_la_zero2 / V_EXAG

                Dim Elevatia1 As String

                If Point_y_min.Y < y01 Then
                    Elevatia1 = "ELEVATION: " & (Get_String_Rounded(Elevatia_cunoscuta - Distanta_de_la_zero1, 1)).ToString
                Else
                    Elevatia1 = "ELEVATION: " & (Get_String_Rounded(Elevatia_cunoscuta + Distanta_de_la_zero1, 1)).ToString
                End If

                Dim Elevatia2 As String

                If Point_y_max.Y < y01 Then
                    Elevatia2 = "ELEVATION: " & (Get_String_Rounded(Elevatia_cunoscuta - Distanta_de_la_zero2, 1)).ToString
                Else
                    Elevatia2 = "ELEVATION: " & (Get_String_Rounded(Elevatia_cunoscuta + Distanta_de_la_zero2, 1)).ToString
                End If



                Dim Mleader1 As New MLeader
                Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_y_min.TransformBy(Curent_UCS), "LOW POINT " & vbCrLf & Elevatia1, 2.5, 2.5, 2, 10, 10)
                Mleader1.Layer = "NO PLOT"

                Dim Mleader2 As New MLeader
                Mleader2 = Creaza_Mleader_nou_fara_UCS_transform(Point_y_max.TransformBy(Curent_UCS), "HIGH POINT " & vbCrLf & Elevatia2, 2.5, 2.5, 2, 10, 10)
                Mleader2.Layer = "NO PLOT"





                Dim Colectie_nume_atr As New Specialized.StringCollection
                Colectie_nume_atr.Add("LOW_HIGH")
                Colectie_nume_atr.Add("TEST_SECTION")
                Colectie_nume_atr.Add("ELEVATION1")
                Colectie_nume_atr.Add("CHAINAGE")

                Dim Colectie_valori_atr_low As New Specialized.StringCollection
                Colectie_valori_atr_low.Add("LOW POINT")
                Colectie_valori_atr_low.Add("TEST SECTION " & nR_TEST_SECTION)
                Colectie_valori_atr_low.Add(Elevatia1)


                Dim Chainage_min As Double
                Dim Chainage_min_text As String

                Dim Chainage_max As Double
                Dim Chainage_max_text As String

                Chainage_min = Chainage_00 + FactorLR * (Point_y_min.X - Point0.Value.X) / HExageration
                Chainage_max = Chainage_00 + FactorLR * (Point_y_max.X - Point0.Value.X) / HExageration
                Chainage_min_text = Get_chainage_from_double(Chainage_min, 1)
                Chainage_max_text = Get_chainage_from_double(Chainage_max, 1)


                Colectie_valori_atr_low.Add(Chainage_min_text)

                InsertBlock_with_multiple_atributes("TEST_SECT_LOW_HIGH_ALIGN.dwg", "LOW1", Point_y_min, 1, BTrecord, Nume_layer, Colectie_nume_atr, Colectie_valori_atr_low)

                Dim Colectie_valori_atr_high As New Specialized.StringCollection
                Colectie_valori_atr_high.Add("HIGH POINT")
                Colectie_valori_atr_high.Add("TEST SECTION " & nR_TEST_SECTION)
                Colectie_valori_atr_high.Add(Elevatia2)
                Colectie_valori_atr_high.Add(Chainage_max_text)

                InsertBlock_with_multiple_atributes("TEST_SECT_LOW_HIGH_ALIGN.dwg", "HIGH1", Point_y_max, 1, BTrecord, Nume_layer, Colectie_nume_atr, Colectie_valori_atr_high)




                Trans1.Commit()
            End Using ' asta e de la trans1
            nR_TEST_SECTION = nR_TEST_SECTION + 1
            GoTo 123


            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception

            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("PPL_ALW2XL")>
    Public Sub Show_ALIGNMENT_W2XL_FORM()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Alignment_w2XL_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Alignment_w2XL_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_TABLE")>
    Public Sub Show_TABLE_TO_EXCEL_FORM()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Table_to_excel_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Table_to_excel_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("PPL_TEXTx_xl")>
    Public Sub Show_TEXT_2_EXCEL_FORM()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is text_2_excel_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New text_2_excel_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("PPL_WKBAND")>
    Public Sub Show_WORKSPACE_BAND_FORM()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Workspace_band_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Workspace_band_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_txt2block")>
    Public Sub Show_text_2_attrib_FORM()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Text_2_attributes_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Text_2_attributes_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_VIEWPORT2POLY")>
    Public Sub VIEWPORT2POLY()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Viewport_to_poly_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Viewport_to_poly_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("PPL_ALBAND")>
    Public Sub Show_ALIGNMENT_BAND_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Alignment_engineering_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Alignment_engineering_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("T2M", CommandFlags.UsePickSet)>
    Public Sub Creeaza_multiple_mtext_objects()
        If isSECURE() = False Then Exit Sub
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select Text objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        For i = 1 To Rezultat1.Value.Count

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i - 1)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Dim Text1 As Autodesk.AutoCAD.DatabaseServices.DBText = Ent1
                                Dim Width1 As Double = Text1.WidthFactor

                                Text1.UpgradeOpen()
                                Using Mtext1 As New MText
                                    Mtext1.BackgroundFill = False
                                    Mtext1.Contents = "{\T" & Width1 & ";" & Text1.TextString & "}"


                                    Mtext1.Layer = Text1.Layer
                                    If Not Text1.AlignmentPoint = New Point3d(0, 0, 0) Then
                                        Mtext1.Location = Text1.AlignmentPoint
                                    Else
                                        Mtext1.Location = Text1.Position
                                    End If
                                    Mtext1.TextHeight = Text1.Height

                                    Mtext1.LineWeight = Text1.LineWeight
                                    Mtext1.TextStyleId = Text1.TextStyleId

                                    Select Case Text1.HorizontalMode
                                        Case TextHorizontalMode.TextLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomLeft
                                        Case TextHorizontalMode.TextMid
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                        Case TextHorizontalMode.TextCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                        Case TextHorizontalMode.TextRight
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomRight
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleRight
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopRight
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomRight

                                        Case Else
                                            Mtext1.Attachment = AttachmentPoint.BottomLeft
                                    End Select

                                    Mtext1.Rotation = Text1.Rotation
                                    Mtext1.ColorIndex = Text1.ColorIndex
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                    Text1.Erase()
                                End Using
                            End If


                        Next
                        Editor1.Regen()
                        Trans1.Commit()
                    End Using ' asta e de  la trans

                Else
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
        Catch ex As Exception
            'Exit Sub
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    <CommandMethod("PPLT2M1", CommandFlags.UsePickSet)>
    Public Sub Creeaza_multiple_mtext_objects_without_alignment_point()
        If isSECURE() = False Then Exit Sub
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select Text objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        For i = 1 To Rezultat1.Value.Count

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i - 1)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Dim Text1 As Autodesk.AutoCAD.DatabaseServices.DBText = Ent1
                                Dim Width1 As Double = Text1.WidthFactor

                                Text1.UpgradeOpen()
                                Using Mtext1 As New MText
                                    Mtext1.BackgroundFill = False
                                    If Width1 <> 1 Then
                                        Mtext1.Contents = "{\T" & Width1 & ";" & Text1.TextString & "}"
                                    Else
                                        Mtext1.Contents = Text1.TextString
                                    End If



                                    Mtext1.Layer = Text1.Layer

                                    Mtext1.Location = Text1.Position

                                    Mtext1.TextHeight = Text1.Height

                                    Mtext1.LineWeight = Text1.LineWeight
                                    Mtext1.TextStyleId = Text1.TextStyleId

                                    Select Case Text1.HorizontalMode
                                        Case TextHorizontalMode.TextLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomLeft
                                        Case TextHorizontalMode.TextMid
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopMid
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                        Case TextHorizontalMode.TextCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopMid
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                        Case TextHorizontalMode.TextRight
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomRight
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleRight
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopRight
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomRight

                                        Case Else
                                            Mtext1.Attachment = AttachmentPoint.BottomLeft
                                    End Select

                                    Mtext1.Rotation = Text1.Rotation
                                    Mtext1.ColorIndex = Text1.ColorIndex
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                    Text1.Erase()
                                End Using
                            End If


                        Next
                        Editor1.Regen()
                        Trans1.Commit()
                    End Using ' asta e de  la trans

                Else
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
        Catch ex As Exception
            'Exit Sub
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub

    <CommandMethod("convert2mtxt", CommandFlags.UsePickSet)>
    Public Sub Cretes_multiple_mtext_objects_from_current_drawing()
        If isSECURE() = False Then Exit Sub
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Sel_result As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Sel_result = Editor1.SelectImplied

            If Sel_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select Text objects:"

                Object_Prompt.SingleOnly = False
                Sel_result = Editor1.GetSelection(Object_Prompt)
            End If

            If Sel_result.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Sel_result) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        For i = 0 To Sel_result.Value.Count - 1

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Sel_result.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Dim Text1 As Autodesk.AutoCAD.DatabaseServices.DBText = Ent1
                                Text1.UpgradeOpen()
                                Using Mtext1 As New MText
                                    Mtext1.BackgroundFill = False

                                    If Text1.WidthFactor = 1 Then
                                        Mtext1.Contents = Text1.TextString
                                    Else
                                        Mtext1.Contents = "{\W" & Text1.WidthFactor.ToString & ";" & Text1.TextString & "}"
                                    End If




                                    Mtext1.Layer = Text1.Layer
                                    If Not Text1.AlignmentPoint = New Point3d(0, 0, 0) Then
                                        Mtext1.Location = Text1.AlignmentPoint
                                    Else
                                        Mtext1.Location = Text1.Position
                                    End If
                                    Mtext1.TextHeight = Text1.Height

                                    Mtext1.LineWeight = Text1.LineWeight
                                    Mtext1.TextStyleId = Text1.TextStyleId

                                    Select Case Text1.HorizontalMode
                                        Case TextHorizontalMode.TextLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopLeft
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomLeft
                                        Case TextHorizontalMode.TextMid
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopMid
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                        Case TextHorizontalMode.TextCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopMid
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomCenter
                                        Case TextHorizontalMode.TextRight
                                            If Text1.VerticalMode = TextVerticalMode.TextBottom Then Mtext1.Attachment = AttachmentPoint.BottomRight
                                            If Text1.VerticalMode = TextVerticalMode.TextVerticalMid Then Mtext1.Attachment = AttachmentPoint.MiddleRight
                                            If Text1.VerticalMode = TextVerticalMode.TextTop Then Mtext1.Attachment = AttachmentPoint.TopRight
                                            If Text1.VerticalMode = TextVerticalMode.TextBase Then Mtext1.Attachment = AttachmentPoint.BottomRight

                                        Case Else
                                            Mtext1.Attachment = AttachmentPoint.BottomLeft
                                    End Select

                                    Mtext1.Rotation = Text1.Rotation
                                    Mtext1.ColorIndex = Text1.ColorIndex
                                    BTrecord.AppendEntity(Mtext1)
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                    Text1.Erase()
                                End Using
                            End If


                        Next
                        Editor1.Regen()
                        Trans1.Commit()
                    End Using ' asta e de  la trans

                Else
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
        Catch ex As Exception
            'Exit Sub
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub


    <CommandMethod("PPL_Simpson")>
    Public Sub Show_heavy_wall_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is heavy_wall_csf_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New heavy_wall_csf_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("PPL_MAT_WRK")>
    Public Sub Show_Alignment_material_worksheet_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Alignment_material_worksheet_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Alignment_material_worksheet_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("ppl_cl", CommandFlags.UsePickSet)>
    Public Sub create_cl_of_2_polys()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 2 polylines:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Not Rezultat1.Value.Count = 2 Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat1.Value.Item(1)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline And TypeOf Ent2 Is Polyline Then
                            Dim Poly1 As Polyline = Ent1
                            Dim Poly2 As Polyline = Ent2

                            Dim data_table1 As New System.Data.DataTable
                            data_table1.Columns.Add("X", GetType(Double))
                            data_table1.Columns.Add("Y", GetType(Double))
                            Dim data_table2 As New System.Data.DataTable
                            data_table2.Columns.Add("X", GetType(Double))
                            data_table2.Columns.Add("Y", GetType(Double))
                            Dim index1 As Double = 0
                            Dim index2 As Double = 0

                            For i = 0 To Poly1.NumberOfVertices - 1
                                data_table1.Rows.Add()
                                data_table1.Rows(index1).Item("X") = Poly1.GetPointAtParameter(i).X
                                data_table1.Rows(index1).Item("Y") = Poly1.GetPointAtParameter(i).Y
                                index1 = index1 + 1
                            Next

                            Dim Pt1 As Point3d
                            Pt1 = Poly1.GetClosestPointTo(Poly2.StartPoint, Vector3d.ZAxis, False)
                            Dim param1 As Double = Poly1.GetParameterAtPoint(Pt1)
                            Dim Pt2 As Point3d
                            Pt2 = Poly1.GetClosestPointTo(Poly2.EndPoint, Vector3d.ZAxis, False)
                            Dim param2 As Double = Poly1.GetParameterAtPoint(Pt2)

                            If param2 < param1 Then
                                For i = Poly2.NumberOfVertices - 1 To 0 Step -1
                                    data_table2.Rows.Add()
                                    data_table2.Rows(index2).Item("X") = Poly2.GetPointAtParameter(i).X
                                    data_table2.Rows(index2).Item("Y") = Poly2.GetPointAtParameter(i).Y
                                    index2 = index2 + 1
                                Next
                            Else
                                For i = 0 To Poly2.NumberOfVertices - 1
                                    data_table2.Rows.Add()
                                    data_table2.Rows(index2).Item("X") = Poly2.GetPointAtParameter(i).X
                                    data_table2.Rows(index2).Item("Y") = Poly2.GetPointAtParameter(i).Y
                                    index2 = index2 + 1
                                Next
                            End If




                            Dim nr_vertice As Double
                            If data_table1.Rows.Count < data_table2.Rows.Count Then
                                nr_vertice = data_table1.Rows.Count
                            Else
                                nr_vertice = data_table2.Rows.Count
                            End If

                            Dim poly_cl As New Polyline
                            For i = 0 To nr_vertice - 1
                                poly_cl.AddVertexAt(i, New Point2d(0.5 * data_table1.Rows(i).Item("X") + 0.5 * data_table2.Rows(i).Item("X"), 0.5 * data_table1.Rows(i).Item("Y") + 0.5 * data_table2.Rows(i).Item("Y")), 0, 0, 0)
                            Next
                            BTrecord.AppendEntity(poly_cl)
                            Trans1.AddNewlyCreatedDBObject(poly_cl, True)


                            Trans1.Commit()
                            Editor1.Regen()
                        End If

                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try

        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
    End Sub


    <CommandMethod("PPL_PI_BREAK")>
    Public Sub Show_split_deflection_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Split_deflection_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Split_deflection_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_split_Bend")>
    Public Sub Show_split_deflection__for_stress_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Split_deflection_for_stress_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Split_deflection_for_stress_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_txt2mtxt_dialog")>
    Public Sub Show_txt2mtxt_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is text2mtext_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New text2mtext_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("w2xlref")>
    Public Sub write_poly_to_excel_with_reference_polyline()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Editor1.SetImpliedSelection(Empty_array)
        Try


            Dim Rezultat0 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult


            Dim Object_Prompt0 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt0.MessageForAdding = vbLf & "Select reference polyline:"

            Object_Prompt0.SingleOnly = True
            Rezultat0 = Editor1.GetSelection(Object_Prompt0)




            If Rezultat0.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select 3D POLYLINES:"

            Object_Prompt.SingleOnly = False
            Rezultat1 = Editor1.GetSelection(Object_Prompt)




            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            Dim RowXL As Integer = 2
            W1 = get_new_worksheet_from_Excel()


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        Dim Obj0 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj0 = Rezultat0.Value.Item(0)
                        Dim Ent0 As Entity
                        Ent0 = Trans1.GetObject(Obj0.ObjectId, OpenMode.ForRead)
                        Dim Poly2D_ref As New Polyline
                        Dim Poly3D_ref As Polyline3d

                        If TypeOf Ent0 Is Polyline3d Then
                            Poly3D_ref = Ent0



                            Dim Index2d0 As Double = 0
                            For Each ObjId As ObjectId In Poly3D_ref
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D_ref.AddVertexAt(Index2d0, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d0 = Index2d0 + 1
                            Next

                            For i = 0 To Rezultat1.Value.Count - 1

                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent1 Is Polyline3d Then
                                    Dim Poly3D As Polyline3d = Ent1
                                    W1.Range("A" & 1).Value = "REFERENCE STA"
                                    W1.Range("B" & 1).Value = "GRID STA"
                                    W1.Range("C" & 1).Value = "X"
                                    W1.Range("D" & 1).Value = "Y"
                                    W1.Range("E" & 1).Value = "Z"


                                    Dim Poly2D As New Polyline
                                    Dim Index2d As Double = 0
                                    For Each ObjId As ObjectId In Poly3D
                                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                        Index2d = Index2d + 1
                                    Next

                                    For j = 0 To Poly2D.NumberOfVertices - 1

                                        Dim pti As New Point3d(Poly3D.GetPointAtParameter(j).X, Poly3D.GetPointAtParameter(j).Y, 0)
                                        Dim pt_on_2d As Point3d = Poly2D_ref.GetClosestPointTo(pti, Vector3d.ZAxis, False)
                                        Dim param1 As Double = Poly2D_ref.GetParameterAtPoint(pt_on_2d)

                                        W1.Range("A" & RowXL).Value = Round(Poly3D_ref.GetDistanceAtParameter(param1), 3)
                                        W1.Range("B" & RowXL).Value = Poly3D.GetDistanceAtParameter(j)
                                        W1.Range("C" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).X, 3)
                                        W1.Range("D" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Y, 3)
                                        W1.Range("E" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Z, 3)

                                        RowXL = RowXL + 1

                                    Next
                                End If
                            Next

                        End If
                    End Using

                Else

                    Exit Sub
                End If
            End If




        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("ppl_CONVERT2D3D", CommandFlags.UsePickSet)>
    Public Sub convert_2Dpoly_to_3d()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select 3D polyline:"

            Object_Prompt2.SingleOnly = True
            Rezultat2 = Editor1.GetSelection(Object_Prompt2)

            If Not Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                        Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj2 = Rezultat2.Value.Item(0)
                        Dim Ent2 As Entity
                        Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Polyline And TypeOf Ent2 Is Polyline3d Then
                            Dim Poly2D As Polyline = Ent1
                            Dim Poly3D As Polyline3d = Ent2
                            Dim Poly2D_3D As New Polyline
                            Dim Index2d As Double = 0
                            For Each ObjId As ObjectId In Poly3D
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D_3D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d = Index2d + 1
                            Next
                            Poly2D_3D.Elevation = Poly2D.Elevation
                            Dim Vertex_from_3D_start As Double = Poly2D_3D.GetParameterAtPoint(Poly2D_3D.GetClosestPointTo(Poly2D.StartPoint, Vector3d.ZAxis, False))
                            Dim Vertex_from_3D_end As Double = Poly2D_3D.GetParameterAtPoint(Poly2D_3D.GetClosestPointTo(Poly2D.EndPoint, Vector3d.ZAxis, False))
                            If Vertex_from_3D_end < Vertex_from_3D_start Then
                                Dim temp As Double = Vertex_from_3D_end
                                Vertex_from_3D_end = Vertex_from_3D_start
                                Vertex_from_3D_start = temp
                            End If

                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("B" & 1).Value = "X"
                            W1.Range("C" & 1).Value = "Y"
                            W1.Range("D" & 1).Value = "Z"

                            For j = 0 To Poly2D.NumberOfVertices - 1

                                W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                W1.Range("B" & RowXL).Value = Round(Poly2D.GetPoint2dAt(j).X, 3)
                                W1.Range("C" & RowXL).Value = Round(Poly2D.GetPoint2dAt(j).Y, 3)
                                W1.Range("D" & RowXL).Value = Round(Poly3D.GetPointAtParameter(Poly2D_3D.GetParameterAtPoint(Poly2D_3D.GetClosestPointTo(Poly2D.GetPointAtParameter(j), Vector3d.ZAxis, False))).Z, 3)
                                If j > 0 And j < Poly2D.NumberOfVertices - 1 Then


                                    Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D.NumberOfVertices - 1 Then
                                                vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                    Dim Mleader1 As New MLeader
                                    Dim LT_RT As String = ""


                                    If Bearing1 < 180 Then
                                        If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                            LT_RT = " LT"
                                        Else
                                            LT_RT = " RT"
                                        End If
                                    Else
                                        If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                            LT_RT = " RT"
                                        Else
                                            LT_RT = " LT"
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

                                    Dim continut_mleader As String = AngleDMS & LT_RT

                                    W1.Range("E" & RowXL).Value = continut_mleader







                                End If
                                RowXL = RowXL + 1

                            Next


                            For i = Ceiling(Vertex_from_3D_start) To Floor(Vertex_from_3D_end)
                                Dim Point_on_3D As New Point3d
                                Point_on_3D = Poly3D.GetPointAtParameter(i)
                                Dim Point_on_2D As New Point3d
                                Point_on_2D = Poly2D.GetClosestPointTo(New Point3d(Point_on_3D.X, Point_on_3D.Y, Poly2D_3D.Elevation), Vector3d.ZAxis, False)
                                W1.Range("A" & RowXL).Value = Poly2D.GetDistAtPoint(Point_on_2D)
                                W1.Range("B" & RowXL).Value = Round(Point_on_2D.X, 3)
                                W1.Range("C" & RowXL).Value = Round(Point_on_2D.Y, 3)
                                W1.Range("D" & RowXL).Value = Round(Point_on_3D.Z, 3)
                                RowXL = RowXL + 1
                            Next

                        End If








                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub



    <CommandMethod("PPL_pipe_length")>
    Public Sub Show_pipe_length_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Length_of_pipe_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Length_of_pipe_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    <CommandMethod("deflection_from_3D", CommandFlags.UsePickSet)>
    Public Sub DEFLECTION_FROM_3D_with_seconds_all_defl()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 3D polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline3d Then
                            Dim Poly2D As New Polyline
                            Dim Poly3D As Polyline3d = Ent1

                            Dim Index2d As Double = 0
                            For Each ObjId As ObjectId In Poly3D
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d = Index2d + 1
                            Next
                            Poly2D.Elevation = 0



                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "GRID DISTANCE"
                            W1.Range("B" & 1).Value = "X"
                            W1.Range("C" & 1).Value = "Y"
                            W1.Range("D" & 1).Value = "Z"
                            W1.Range("E" & 1).Value = "DEFLECTION"
                            W1.Range("F" & 1).Value = "DEFLECTION DD"

                            For j = 0 To Poly2D.NumberOfVertices - 1



                                If j > 0 And j < Poly2D.NumberOfVertices - 1 Then

                                    Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D.NumberOfVertices - 1 Then
                                                vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI

                                    Dim Mleader1 As New MLeader
                                    Dim LT_RT As String = ""


                                    If Bearing1 < 180 Then
                                        If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                            LT_RT = " LT"
                                        Else
                                            LT_RT = " RT"
                                        End If
                                    Else
                                        If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                            LT_RT = " RT"
                                        Else
                                            LT_RT = " LT"
                                        End If
                                    End If



                                    Dim AngleDMS As String = Floor(angle1) & "°"
                                    Dim Minute1 As Double = Floor((angle1 - Floor(angle1)) * 60)

                                    Dim Minute As String = Floor((angle1 - Floor(angle1)) * 60) & "'"

                                    If Minute1 = 60 Then
                                        AngleDMS = Floor(angle1 + 1) & "°"
                                        Minute = "00'"
                                    End If
                                    Dim Second As String = Round((angle1 - Floor(angle1) - Minute1 / 60) * 3600, 0) & Chr(34)

                                    If Len(Minute) = 2 Then Minute = "0" & Minute
                                    If Len(Second) = 2 Then Second = "0" & Second
                                    If Round((angle1 - Floor(angle1) - Minute1 / 60) * 3600, 0) = 60 Then
                                        Minute1 = Minute1 + 1
                                        If Minute1 = 60 Then
                                            AngleDMS = Floor(angle1 + 1) & "°"
                                            Minute = "00'"
                                        End If
                                        Second = "00" & Chr(34)
                                    End If


                                    AngleDMS = AngleDMS & Minute & Second



                                    Dim continut_mleader As String = AngleDMS & LT_RT
                                    If angle1 >= 0.5 Then
                                        W1.Range("A" & RowXL).Value = Poly3D.GetDistanceAtParameter(j)
                                        W1.Range("B" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).X, 3)
                                        W1.Range("C" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Y, 3)
                                        W1.Range("D" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Z, 3)
                                        W1.Range("E" & RowXL).Value = continut_mleader
                                        W1.Range("F" & RowXL).Value = Round(angle1, 3)

                                        Creaza_Mleader_nou_fara_UCS_transform(Poly3D.GetPointAtParameter(j), "X = " & Round(Poly3D.GetPointAtParameter(j).X, 3) & vbCrLf &
                                                                              "Y = " & Round(Poly3D.GetPointAtParameter(j).Y, 3) & vbCrLf &
                                                                              "Z = " & Round(Poly3D.GetPointAtParameter(j).Z, 3) & vbCrLf &
                                                                              continut_mleader, 2.5, 2.5, 2.5, 5, 10)

                                        RowXL = RowXL + 1
                                    End If




                                Else
                                    W1.Range("A" & RowXL).Value = Poly3D.GetDistanceAtParameter(j)
                                    W1.Range("B" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).X, 3)
                                    W1.Range("C" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Y, 3)
                                    W1.Range("D" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Z, 3)

                                    RowXL = RowXL + 1

                                End If
                            Next

                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    <CommandMethod("ppl_DEFL_FROM_3d", CommandFlags.UsePickSet)>
    Public Sub DEFLECTION_FROM_3D()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 3D polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline3d Then
                            Dim Poly2D As New Polyline
                            Dim Poly3D As Polyline3d = Ent1

                            Dim Index2d As Double = 0
                            For Each ObjId As ObjectId In Poly3D
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d = Index2d + 1
                            Next
                            Poly2D.Elevation = 0



                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "GRID DISTANCE"
                            W1.Range("B" & 1).Value = "X"
                            W1.Range("C" & 1).Value = "Y"
                            W1.Range("D" & 1).Value = "Z"
                            W1.Range("E" & 1).Value = "DEFLECTION"
                            W1.Range("F" & 1).Value = "DEFLECTION DD"

                            For j = 0 To Poly2D.NumberOfVertices - 1



                                If j > 0 And j < Poly2D.NumberOfVertices - 1 Then

                                    Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D.NumberOfVertices - 1 Then
                                                vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                    Dim Mleader1 As New MLeader
                                    Dim LT_RT As String = ""


                                    If Bearing1 < 180 Then
                                        If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                            LT_RT = " LT"
                                        Else
                                            LT_RT = " RT"
                                        End If
                                    Else
                                        If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                            LT_RT = " RT"
                                        Else
                                            LT_RT = " LT"
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

                                    Dim continut_mleader As String = AngleDMS & LT_RT
                                    If angle1 >= 0.5 Then

                                        W1.Range("A" & RowXL).Value = Poly3D.GetDistanceAtParameter(j)
                                        W1.Range("B" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).X, 3)
                                        W1.Range("C" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Y, 3)
                                        W1.Range("D" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Z, 3)
                                        W1.Range("E" & RowXL).Value = continut_mleader
                                        W1.Range("F" & RowXL).Value = Round(angle1, 3)

                                        Creaza_Mleader_nou_fara_UCS_transform(Poly3D.GetPointAtParameter(j), "X = " & Round(Poly3D.GetPointAtParameter(j).X, 3) & vbCrLf &
                                                                              "Y = " & Round(Poly3D.GetPointAtParameter(j).Y, 3) & vbCrLf &
                                                                              "Z = " & Round(Poly3D.GetPointAtParameter(j).Z, 3) & vbCrLf &
                                                                              continut_mleader, 2.5, 2.5, 2.5, 5, 10)

                                        RowXL = RowXL + 1

                                    End If
                                Else
                                    W1.Range("A" & RowXL).Value = Poly3D.GetDistanceAtParameter(j)
                                    W1.Range("B" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).X, 3)
                                    W1.Range("C" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Y, 3)
                                    W1.Range("D" & RowXL).Value = Round(Poly3D.GetPointAtParameter(j).Z, 3)

                                    RowXL = RowXL + 1

                                End If
                            Next

                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub




    <CommandMethod("ppl_DEFL_FROM_2d", CommandFlags.UsePickSet)>
    Public Sub DEFLECTION_FROM_2D()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 2D polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline Then
                            Dim Poly2D As Polyline = Ent1

                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "GRID DISTANCE"
                            W1.Range("B" & 1).Value = "X"
                            W1.Range("C" & 1).Value = "Y"
                            W1.Range("D" & 1).Value = "Z"
                            W1.Range("E" & 1).Value = "DEFLECTION"
                            W1.Range("F" & 1).Value = "DEFLECTION DD"

                            For j = 0 To Poly2D.NumberOfVertices - 1



                                If j > 0 And j < Poly2D.NumberOfVertices - 1 Then


                                    Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D.NumberOfVertices - 1 Then
                                                vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                    Dim Mleader1 As New MLeader
                                    Dim LT_RT As String = ""


                                    If Bearing1 < 180 Then
                                        If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                            LT_RT = " LT"
                                        Else
                                            LT_RT = " RT"
                                        End If
                                    Else
                                        If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                            LT_RT = " RT"
                                        Else
                                            LT_RT = " LT"
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

                                    Dim continut_mleader As String = AngleDMS & LT_RT

                                    If angle1 >= 0.5 Then

                                        W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                        W1.Range("B" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                        W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                        W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)
                                        W1.Range("E" & RowXL).Value = continut_mleader
                                        W1.Range("F" & RowXL).Value = Round(angle1, 3)

                                        Creaza_Mleader_nou_fara_UCS_transform(Poly2D.GetPointAtParameter(j), "X = " & Round(Poly2D.GetPointAtParameter(j).X, 3) & vbCrLf &
                                                                              "Y = " & Round(Poly2D.GetPointAtParameter(j).Y, 3) & vbCrLf &
                                                                              "Z = " & Round(Poly2D.GetPointAtParameter(j).Z, 3) & vbCrLf &
                                                                              continut_mleader, 2.5, 2.5, 2.5, 5, 10)


                                        RowXL = RowXL + 1

                                    End If
                                Else
                                    W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                    W1.Range("B" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                    W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                    W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)

                                    RowXL = RowXL + 1
                                End If
                            Next
                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("ppl_DEFL_FROM_2d0", CommandFlags.UsePickSet)>
    Public Sub DEFLECTION_FROM_2D_with_seconds()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 2D polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline Then
                            Dim Poly2D As Polyline = Ent1

                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "GRID DISTANCE"
                            W1.Range("B" & 1).Value = "X"
                            W1.Range("C" & 1).Value = "Y"
                            W1.Range("D" & 1).Value = "Z"
                            W1.Range("E" & 1).Value = "DEFLECTION"
                            W1.Range("F" & 1).Value = "DEFLECTION DD"

                            For j = 0 To Poly2D.NumberOfVertices - 1



                                If j > 0 And j < Poly2D.NumberOfVertices - 1 Then


                                    Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D.NumberOfVertices - 1 Then
                                                vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                    Dim Mleader1 As New MLeader
                                    Dim LT_RT As String = ""


                                    If Bearing1 < 180 Then
                                        If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                            LT_RT = " LT"
                                        Else
                                            LT_RT = " RT"
                                        End If
                                    Else
                                        If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                            LT_RT = " RT"
                                        Else
                                            LT_RT = " LT"
                                        End If

                                    End If

                                    Dim AngleDMS As String = Floor(angle1) & "°"
                                    Dim Minute1 As Double = Floor((angle1 - Floor(angle1)) * 60)

                                    Dim Minute As String = Floor((angle1 - Floor(angle1)) * 60) & "'"

                                    If Minute1 = 60 Then
                                        AngleDMS = Floor(angle1 + 1) & "°"
                                        Minute = "00'"
                                    End If
                                    Dim Second As String = Round((angle1 - Floor(angle1) - Minute1 / 60) * 3600, 0) & Chr(34)

                                    If Len(Minute) = 2 Then Minute = "0" & Minute
                                    If Len(Second) = 2 Then Second = "0" & Second
                                    If Round((angle1 - Floor(angle1) - Minute1 / 60) * 3600, 0) = 60 Then
                                        Minute1 = Minute1 + 1
                                        If Minute1 = 60 Then
                                            AngleDMS = Floor(angle1 + 1) & "°"
                                            Minute = "00'"
                                        End If
                                        Second = "00" & Chr(34)
                                    End If


                                    AngleDMS = AngleDMS & Minute & Second

                                    Dim continut_mleader As String = AngleDMS & LT_RT

                                    If angle1 >= 0.5 Then

                                        W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                        W1.Range("B" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                        W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                        W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)
                                        W1.Range("E" & RowXL).Value = continut_mleader
                                        W1.Range("F" & RowXL).Value = Round(angle1, 3)

                                        Creaza_Mleader_nou_fara_UCS_transform(Poly2D.GetPointAtParameter(j), "X = " & Round(Poly2D.GetPointAtParameter(j).X, 3) & vbCrLf &
                                                                              "Y = " & Round(Poly2D.GetPointAtParameter(j).Y, 3) & vbCrLf &
                                                                              "Z = " & Round(Poly2D.GetPointAtParameter(j).Z, 3) & vbCrLf &
                                                                              continut_mleader, 2.5, 2.5, 2.5, 5, 10)


                                        RowXL = RowXL + 1

                                    End If
                                Else
                                    W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                    W1.Range("B" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                    W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                    W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)

                                    RowXL = RowXL + 1
                                End If
                            Next
                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("PPL_DEFL_VERT_3D", CommandFlags.UsePickSet)>
    Public Sub VERTICAL_DEFLECTION_FROM_3D()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied
            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 3D polyline:"
                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2




                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline3d Then
                            Dim Data_table As New System.Data.DataTable
                            Data_table.Columns.Add("X", GetType(Double))
                            Data_table.Columns.Add("Y", GetType(Double))
                            Data_table.Columns.Add("Z", GetType(Double))
                            Data_table.Columns.Add("STATION", GetType(Double))
                            Data_table.Columns.Add("VERTDEFL", GetType(Double))

                            Dim Poly3D As Polyline3d = Ent1



                            For i = 0 To Poly3D.EndParam Step 1
                                Data_table.Rows.Add()
                                Dim x, y, z As Double
                                x = Poly3D.GetPointAtParameter(i).X
                                y = Poly3D.GetPointAtParameter(i).Y
                                z = Poly3D.GetPointAtParameter(i).Z
                                Data_table.Rows(i).Item("X") = x
                                Data_table.Rows(i).Item("Y") = y
                                Data_table.Rows(i).Item("Z") = z
                                Data_table.Rows(i).Item("STATION") = Poly3D.GetDistAtPoint(Poly3D.GetPointAtParameter(i))



                                If i = 0 Then
                                    Data_table.Rows(i).Item("VERTDEFL") = 0

                                ElseIf i = Poly3D.EndParam Then
                                    Data_table.Rows(i).Item("VERTDEFL") = 0
                                Else
                                    Dim Dist2D_a As Double
                                    Dim Dist2D_b As Double
                                    Dim DeltaZ_a As Double
                                    Dim DeltaZ_b As Double
                                    Dim Alpha1 As Double
                                    Dim Alpha2 As Double
                                    Dim Vdefl As Double

                                    Dist2D_a = ((Poly3D.GetPointAtParameter(i - 1).X - x) ^ 2 + (Poly3D.GetPointAtParameter(i - 1).Y - y) ^ 2) ^ 0.5
                                    DeltaZ_a = z - Poly3D.GetPointAtParameter(i - 1).Z
                                    Alpha1 = Atan(Abs(DeltaZ_a) / Dist2D_a)

                                    Dim A1 As Double = Alpha1 * 180 / PI

                                    Dist2D_b = ((Poly3D.GetPointAtParameter(i + 1).X - x) ^ 2 + (Poly3D.GetPointAtParameter(i + 1).Y - y) ^ 2) ^ 0.5
                                    DeltaZ_b = Poly3D.GetPointAtParameter(i + 1).Z - z
                                    Alpha2 = Atan(Abs(DeltaZ_b) / Dist2D_b)

                                    Dim A2 As Double = Alpha2 * 180 / PI

                                    If DeltaZ_b >= 0 And DeltaZ_a >= 0 Then
                                        Vdefl = Abs(Alpha2 - Alpha1)
                                    ElseIf DeltaZ_b < 0 And DeltaZ_a >= 0 Then
                                        Vdefl = Alpha2 + Alpha1
                                    ElseIf DeltaZ_b >= 0 And DeltaZ_a < 0 Then
                                        Vdefl = Alpha2 + Alpha1
                                    ElseIf DeltaZ_b < 0 And DeltaZ_a < 0 Then
                                        Vdefl = Abs(Alpha2 - Alpha1)
                                    End If

                                    Data_table.Rows(i).Item("VERTDEFL") = Vdefl * 180 / PI
                                End If


                            Next




                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "GRID DISTANCE"
                            W1.Range("B" & 1).Value = "X"
                            W1.Range("C" & 1).Value = "Y"
                            W1.Range("D" & 1).Value = "Z"
                            W1.Range("E" & 1).Value = "DEFLECTION"
                            W1.Range("F" & 1).Value = "DEFLECTION DD"

                            For i = 0 To Data_table.Rows.Count - 1


                                Dim Angle1 As Double = Data_table.Rows(i).Item("VERTDEFL")
                                Dim AngleDMS As String = Floor(Angle1) & "°"
                                Dim Minute As String = Round((Angle1 - Floor(Angle1)) * 60, 0) & "'"
                                If Len(Minute) = 2 Then Minute = "0" & Minute
                                AngleDMS = AngleDMS & Minute & "00" & Chr(34)
                                If Round((Angle1 - Floor(Angle1)) * 60, 0) = 60 Then
                                    AngleDMS = Floor(Angle1 + 1) & "°"
                                    Minute = "00'"
                                End If


                                W1.Range("A" & RowXL).Value = Round(Data_table.Rows(i).Item("STATION"), 3)
                                W1.Range("B" & RowXL).Value = Round(Data_table.Rows(i).Item("X"), 3)
                                W1.Range("C" & RowXL).Value = Round(Data_table.Rows(i).Item("Y"), 3)
                                W1.Range("D" & RowXL).Value = Round(Data_table.Rows(i).Item("Z"), 3)
                                W1.Range("E" & RowXL).Value = AngleDMS
                                W1.Range("F" & RowXL).Value = Round(Angle1, 3)


                                RowXL = RowXL + 1


                            Next

                        End If

                        Trans1.Abort()
                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("ReferenceXL")>
    Public Sub Show_ref_from_xl_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is References_from_excel_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New References_from_excel_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_insert_BLK")>
    Public Sub Show__layout_block_insert_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Block_layout_insert_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Block_layout_insert_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_BLK2XL", CommandFlags.UsePickSet)>
    Public Sub write_block_attributes_to_excel()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        Dim Empty_array() As ObjectId
        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select blocks containing the attributes:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.SetImpliedSelection(Empty_array)
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        W1 = get_new_worksheet_from_Excel()
                        Dim RowXL As Integer = 2

                        For i = 0 To Rezultat1.Value.Count - 1

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            If TypeOf Ent1 Is BlockReference Then
                                Dim Block1 As BlockReference = Ent1
                                If Block1.AttributeCollection.Count > 0 Then
                                    For Each Id1 As ObjectId In Block1.AttributeCollection
                                        Dim Atr1 As AttributeReference
                                        Atr1 = TryCast(Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                        If IsNothing(Atr1) = False Then
                                            Dim Tag1 As String = Atr1.Tag
                                            Dim Val1 As String
                                            If Atr1.IsMTextAttribute = True Then
                                                Val1 = Atr1.MTextAttribute.Contents
                                            Else
                                                Val1 = Atr1.TextString
                                            End If

                                            W1.Range("A" & RowXL).Value = Tag1
                                            If String.IsNullOrEmpty(Val1) = False Then
                                                W1.Range("B" & RowXL).Value = Val1
                                            End If
                                            RowXL = RowXL + 1
                                        End If
                                    Next
                                End If





                            End If










                        Next





                    End Using

                Else
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                    Editor1.SetImpliedSelection(Empty_array)
                    Exit Sub
                End If
            End If


            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    <CommandMethod("PROFILE_VIEWPORT")>
    Public Sub Show_VIEWPORT_ALONG_PROF_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Viewports_along_graph_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Viewports_along_graph_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("ppl_listplot", CommandFlags.UsePickSet)>
    Public Sub list_plot_PDF()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)


                Dim LayoutManager1 As LayoutManager = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current
                Dim Layoutdict As DBDictionary = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                Dim Plot_settings_dict As DBDictionary = Trans1.GetObject(ThisDrawing.Database.PlotSettingsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim Layout1 As Layout = Trans1.GetObject(LayoutManager1.GetLayoutId(LayoutManager1.CurrentLayout), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim Plott_settings As PlotSettings
                If Plot_settings_dict.Contains("B1_PDF") = False Then
                    Dim Plot_info As New Autodesk.AutoCAD.PlottingServices.PlotInfo


                    Plot_info.Layout = Layout1.ObjectId
                    Plott_settings = New PlotSettings(Layout1.ModelType)
                    Plott_settings.CopyFrom(Layout1)
                    Plott_settings.PlotSettingsName = "B1_PDF"
                Else
                    Plott_settings = Plot_settings_dict.GetAt("B1_PDF").GetObject(OpenMode.ForWrite)
                End If


                Try
                    Dim Plot_Settings_Validator As PlotSettingsValidator = PlotSettingsValidator.Current
                    Plot_Settings_Validator.RefreshLists(Plott_settings)
                    Dim MediaList As New Specialized.StringCollection
                    MediaList = Plot_Settings_Validator.GetCanonicalMediaNameList(Plott_settings)

                    For Each STR1 As String In MediaList
                        Editor1.WriteMessage(vbLf & STR1)
                    Next




                Catch es As Autodesk.AutoCAD.Runtime.Exception
                    MsgBox(es.Message)
                End Try






            End Using







            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub
    <CommandMethod("Crrrr")>
    Public Shared Sub CreateOrEditPageSetup()
        ' Get the current document and database, and start a transaction
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim plSets As DBDictionary =
                acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim vStyles As DBDictionary =
                acTrans.GetObject(acCurDb.VisualStyleDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

            Dim acPlSet As PlotSettings
            Dim createNew As Boolean = False

            ' Reference the Layout Manager
            Dim acLayoutMgr As LayoutManager = LayoutManager.Current

            ' Get the current layout and output its name in the Command Line window
            Dim acLayout As Layout =
                acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                  OpenMode.ForRead)

            ' Check to see if the page setup exists
            If plSets.Contains("B1_PDF") = False Then
                createNew = True

                ' Create a new PlotSettings object: 
                '    True - model space, False - named layout
                acPlSet = New PlotSettings(acLayout.ModelType)
                acPlSet.CopyFrom(acLayout)

                acPlSet.PlotSettingsName = "B1_PDF"
                acPlSet.AddToPlotSettingsDictionary(acCurDb)
                acTrans.AddNewlyCreatedDBObject(acPlSet, True)
            Else
                acPlSet = plSets.GetAt("B1_PDF").GetObject(OpenMode.ForWrite)
            End If

            ' Update the PlotSettings object
            Try
                Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current

                ' Set the Plotter and page size
                acPlSetVdr.SetPlotConfigurationName(acPlSet, "Adobe PDF(B1).pc3", "User32767")

                ' Set to plot to the current display
                If acLayout.ModelType = False Then
                    ' Use SetPlotWindowArea with PlotType.Window
                    acPlSetVdr.SetPlotWindowArea(acPlSet, New Extents2d(New Point2d(0.0, 0.0), New Point2d(1000.0, 707.0)))
                    acPlSetVdr.SetPlotType(acPlSet, PlotType.Window)


                Else
                    acPlSetVdr.SetPlotType(acPlSet,
                                           PlotType.Extents)

                    acPlSetVdr.SetPlotCentered(acPlSet, True)
                End If



                ' Use SetPlotViewName with PlotType.View
                'acPlSetVdr.SetPlotViewName(plSet, "MyView")

                ' Set the plot offset
                'acPlSetVdr.SetPlotOrigin(acPlSet,    New Point2d(0, 0))

                ' Set the plot scale
                'acPlSetVdr.SetUseStandardScale(acPlSet, False)

                'acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.StdScale1To1)


                Dim Scale As CustomScale = New CustomScale(1, 1)
                acPlSetVdr.SetCustomPrintScale(acPlSet, Scale)
                acPlSet.ScaleLineweights = False
                acPlSetVdr.SetPlotCentered(acPlSet, True)

                ' Specify if plot styles should be displayed on the layout
                acPlSet.ShowPlotStyles = False

                ' Rebuild plotter, plot style, and canonical media lists 
                ' (must be called before setting the plot style)
                acPlSetVdr.RefreshLists(acPlSet)

                ' Specify the shaded viewport options
                acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed

                acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal

                ' Specify the plot options
                acPlSet.PrintLineweights = True
                acPlSet.PlotTransparency = False
                acPlSet.PlotPlotStyles = True
                acPlSet.DrawViewportsFirst = True

                ' Use only on named layouts - Hide paperspace objects option
                ' plSet.PlotHidden = True

                ' Specify the plot orientation
                acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees090)

                ' Set the plot style
                If acCurDb.PlotStyleMode = True Then
                    acPlSetVdr.SetCurrentStyleSheet(acPlSet, "TC_MONO.ctb")
                Else
                    acPlSetVdr.SetCurrentStyleSheet(acPlSet, "acad.stb")
                End If
                acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Millimeters)
                ' Zoom to show the whole paper
                acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, True)

            Catch es As Autodesk.AutoCAD.Runtime.Exception
                MsgBox(es.Message)
            End Try

            ' Save the changes made
            acTrans.Commit()

            If createNew = True Then
                acPlSet.Dispose()
            End If
        End Using
    End Sub
    <CommandMethod("PPL_rw_builder")>
    Public Sub Show_rw_builder_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is RW_Builder Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New RW_Builder
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    '<CommandMethod("Sort_poly_x")> _
    Public Sub Sort_poly_x()

        If isSECURE() = False Then Exit Sub



        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor





            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                Dim Rezultat_lines_and_polylies As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt2.MessageForAdding = vbLf & "Select the multiple_lines:"
                Object_Prompt2.SingleOnly = False
                Rezultat_lines_and_polylies = Editor1.GetSelection(Object_Prompt2)





                Creaza_layer("NO PLOT", 40, "", False)


                Dim Colectie_puncte As New Point3dCollection
                Dim Table_data1 As New System.Data.DataTable
                Table_data1.Columns.Add("X", GetType(Double))
                Table_data1.Columns.Add("Y", GetType(Double))

                Dim Index_table As Double = 0
                For i = 1 To Rezultat_lines_and_polylies.Value.Count
                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                    Obj1 = Rezultat_lines_and_polylies.Value.Item(i - 1)
                    Dim Ent1 As Entity
                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Line Then
                        Dim Linie1 As Line = Ent1
                        Table_data1.Rows.Add()
                        Table_data1.Rows(Index_table).Item("X") = Linie1.StartPoint.X
                        Table_data1.Rows(Index_table).Item("Y") = Linie1.StartPoint.Y
                        Index_table = Index_table + 1
                        Table_data1.Rows.Add()
                        Table_data1.Rows(Index_table).Item("X") = Linie1.EndPoint.X
                        Table_data1.Rows(Index_table).Item("Y") = Linie1.EndPoint.Y
                        Index_table = Index_table + 1
                    End If
                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        Dim PLinie1 As Polyline = Ent1

                        For j = 0 To PLinie1.NumberOfVertices - 1
                            Table_data1.Rows.Add()
                            Table_data1.Rows(Index_table).Item("X") = PLinie1.GetPoint3dAt(j).X
                            Table_data1.Rows(Index_table).Item("Y") = PLinie1.GetPoint3dAt(j).Y
                            Index_table = Index_table + 1
                        Next

                    End If
                Next
                Dim DataView1 As New DataView(Table_data1)
                DataView1.Sort = "X"

                Dim Poly_graph As New Polyline

                Dim z As Double = 0
                For Each row1 In DataView1
                    Poly_graph.AddVertexAt(z, New Point2d(row1("X"), row1("Y")), 0, 0, 0)
                    z = z + 1
                Next


                Poly_graph.Layer = "NO PLOT"
                BTrecord.AppendEntity(Poly_graph)
                Trans1.AddNewlyCreatedDBObject(Poly_graph, True)








                Trans1.Commit()


            End Using



            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("PPL_chspace")>
    Public Sub Show_styles_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Standard_Styles_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Standard_Styles_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_easement")>
    Public Sub Show_easement_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Easement_builder_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Easement_builder_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("PPL_styles_XL")>
    Public Sub Text_style_to_excel()
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
            W1 = get_new_worksheet_from_Excel()
            Dim Row1 As Integer = 1

            Using lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Text_style_table As Autodesk.AutoCAD.DatabaseServices.TextStyleTable = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    For Each Id1 As ObjectId In Text_style_table
                        Dim style1 As TextStyleTableRecord = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), TextStyleTableRecord)
                        If IsNothing(style1) = False Then
                            Select Case MsgBox("TEXTSTYLE " & style1.Name & "?", MsgBoxStyle.YesNoCancel)
                                Case MsgBoxResult.Yes
                                    W1.Range("A" & Row1).Value = "TEXTSTYLE.NAME"
                                    W1.Range("B" & Row1).Value = style1.Name
                                    W1.Range("A" & Row1 + 1).Value = ".FILENAME"
                                    W1.Range("B" & Row1 + 1).Value = style1.FileName
                                    W1.Range("A" & Row1 + 2).Value = ".TEXTSIZE"
                                    W1.Range("B" & Row1 + 2).Value = style1.TextSize
                                    W1.Range("A" & Row1 + 3).Value = ".ObliquingAngle"
                                    W1.Range("B" & Row1 + 3).Value = style1.ObliquingAngle
                                    W1.Range("A" & Row1 + 4).Value = ".XScale"
                                    W1.Range("B" & Row1 + 4).Value = style1.XScale
                                    Row1 = Row1 + 5
                                Case MsgBoxResult.Cancel
                                    Exit For
                            End Select
                        End If

                    Next

                    Dim Dim_style_table As Autodesk.AutoCAD.DatabaseServices.DimStyleTable = Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    For Each Id1 As ObjectId In Dim_style_table
                        Dim style1 As DimStyleTableRecord = TryCast(Trans1.GetObject(Id1, OpenMode.ForRead), DimStyleTableRecord)
                        If IsNothing(style1) = False Then
                            Select Case MsgBox("DIMSTYLE " & style1.Name & "?", MsgBoxStyle.YesNoCancel)
                                Case MsgBoxResult.Yes
                                    W1.Range("A" & Row1 + 0).Value = "DIMSTYLE Name"
                                    W1.Range("A" & Row1 + 1).Value = "Dimadec"
                                    W1.Range("A" & Row1 + 2).Value = "Dimalt"
                                    W1.Range("A" & Row1 + 3).Value = "Dimaltd"
                                    W1.Range("A" & Row1 + 4).Value = "Dimaltf"
                                    W1.Range("A" & Row1 + 5).Value = "Dimaltrnd"
                                    W1.Range("A" & Row1 + 6).Value = "Dimalttd"
                                    W1.Range("A" & Row1 + 7).Value = "Dimalttz"
                                    W1.Range("A" & Row1 + 8).Value = "Dimaltu"
                                    W1.Range("A" & Row1 + 9).Value = "Dimaltz"
                                    W1.Range("A" & Row1 + 10).Value = "Dimapost"
                                    W1.Range("A" & Row1 + 11).Value = "Dimarcsym"
                                    W1.Range("A" & Row1 + 12).Value = "Dimasz"
                                    W1.Range("A" & Row1 + 13).Value = "Dimatfit"
                                    W1.Range("A" & Row1 + 14).Value = "Dimaunit"
                                    W1.Range("A" & Row1 + 15).Value = "Dimazin"
                                    W1.Range("A" & Row1 + 16).Value = "Dimcen"
                                    W1.Range("A" & Row1 + 17).Value = "Dimclrd"
                                    W1.Range("A" & Row1 + 18).Value = "Dimclre"
                                    W1.Range("A" & Row1 + 19).Value = "Dimclrt"
                                    W1.Range("A" & Row1 + 20).Value = "Dimdec"
                                    W1.Range("A" & Row1 + 21).Value = "Dimdle"
                                    W1.Range("A" & Row1 + 22).Value = "Dimdli"
                                    W1.Range("A" & Row1 + 23).Value = "Dimdsep"
                                    W1.Range("A" & Row1 + 24).Value = "Dimexe"
                                    W1.Range("A" & Row1 + 25).Value = "Dimexo"
                                    W1.Range("A" & Row1 + 26).Value = "Dimfrac"
                                    W1.Range("A" & Row1 + 27).Value = "Dimfxlen"
                                    W1.Range("A" & Row1 + 28).Value = "DimfxlenOn"
                                    W1.Range("A" & Row1 + 29).Value = "Dimgap"
                                    W1.Range("A" & Row1 + 30).Value = "Dimjogang"
                                    W1.Range("A" & Row1 + 31).Value = "Dimjust"
                                    W1.Range("A" & Row1 + 32).Value = "Dimblk"
                                    W1.Range("A" & Row1 + 33).Value = "Dimldrblk"
                                    W1.Range("A" & Row1 + 34).Value = "Dimlfac"
                                    W1.Range("A" & Row1 + 35).Value = "Dimlim"
                                    W1.Range("A" & Row1 + 36).Value = "Dimlunit"
                                    W1.Range("A" & Row1 + 37).Value = "Dimlwd"
                                    W1.Range("A" & Row1 + 38).Value = "Dimlwe"
                                    W1.Range("A" & Row1 + 39).Value = "Dimpost"
                                    W1.Range("A" & Row1 + 40).Value = "Dimrnd"
                                    W1.Range("A" & Row1 + 41).Value = "Dimsah"
                                    W1.Range("A" & Row1 + 42).Value = "Dimscale"
                                    W1.Range("A" & Row1 + 43).Value = "Dimsd1"
                                    W1.Range("A" & Row1 + 44).Value = "Dimsd2"
                                    W1.Range("A" & Row1 + 45).Value = "Dimse1"
                                    W1.Range("A" & Row1 + 46).Value = "Dimse2"
                                    W1.Range("A" & Row1 + 47).Value = "Dimsoxd"
                                    W1.Range("A" & Row1 + 48).Value = "Dimtad"
                                    W1.Range("A" & Row1 + 49).Value = "Dimtdec"
                                    W1.Range("A" & Row1 + 50).Value = "Dimtfac"
                                    W1.Range("A" & Row1 + 51).Value = "Dimtfill"
                                    W1.Range("A" & Row1 + 52).Value = "Dimtfillclr"
                                    W1.Range("A" & Row1 + 53).Value = "Dimtih"
                                    W1.Range("A" & Row1 + 54).Value = "Dimtix"
                                    W1.Range("A" & Row1 + 55).Value = "Dimtm"
                                    W1.Range("A" & Row1 + 56).Value = "Dimtmove"
                                    W1.Range("A" & Row1 + 57).Value = "Dimtofl"
                                    W1.Range("A" & Row1 + 58).Value = "Dimtoh"
                                    W1.Range("A" & Row1 + 59).Value = "Dimtol"
                                    W1.Range("A" & Row1 + 60).Value = "Dimtolj"
                                    W1.Range("A" & Row1 + 61).Value = "Dimtp"
                                    W1.Range("A" & Row1 + 62).Value = "Dimtsz"
                                    W1.Range("A" & Row1 + 63).Value = "Dimtvp"
                                    W1.Range("A" & Row1 + 64).Value = "Dimtxsty"
                                    W1.Range("A" & Row1 + 65).Value = "Dimtxt"
                                    W1.Range("A" & Row1 + 66).Value = "Dimtxtdirection"
                                    W1.Range("A" & Row1 + 67).Value = "Dimtzin"
                                    W1.Range("A" & Row1 + 68).Value = "Dimupt"
                                    W1.Range("A" & Row1 + 69).Value = "Dimzin"




                                    W1.Range("B" & Row1 + 0).Value = style1.Name
                                    W1.Range("B" & Row1 + 1).Value = style1.Dimadec
                                    W1.Range("B" & Row1 + 2).Value = style1.Dimalt
                                    W1.Range("B" & Row1 + 3).Value = style1.Dimaltd
                                    W1.Range("B" & Row1 + 4).Value = style1.Dimaltf
                                    W1.Range("B" & Row1 + 5).Value = style1.Dimaltrnd
                                    W1.Range("B" & Row1 + 6).Value = style1.Dimalttd
                                    W1.Range("B" & Row1 + 7).Value = style1.Dimalttz
                                    W1.Range("B" & Row1 + 8).Value = style1.Dimaltu
                                    W1.Range("B" & Row1 + 9).Value = style1.Dimaltz
                                    W1.Range("B" & Row1 + 10).Value = style1.Dimapost
                                    W1.Range("B" & Row1 + 11).Value = style1.Dimarcsym
                                    W1.Range("B" & Row1 + 12).Value = style1.Dimasz
                                    W1.Range("B" & Row1 + 13).Value = style1.Dimatfit
                                    W1.Range("B" & Row1 + 14).Value = style1.Dimaunit
                                    W1.Range("B" & Row1 + 15).Value = style1.Dimazin
                                    W1.Range("B" & Row1 + 16).Value = style1.Dimcen
                                    W1.Range("B" & Row1 + 17).Value = style1.Dimclrd.ColorIndex
                                    W1.Range("B" & Row1 + 18).Value = style1.Dimclre.ColorIndex
                                    W1.Range("B" & Row1 + 19).Value = style1.Dimclrt.ColorIndex
                                    W1.Range("B" & Row1 + 20).Value = style1.Dimdec
                                    W1.Range("B" & Row1 + 21).Value = style1.Dimdle
                                    W1.Range("B" & Row1 + 22).Value = style1.Dimdli
                                    W1.Range("B" & Row1 + 23).Value = style1.Dimdsep
                                    W1.Range("B" & Row1 + 24).Value = style1.Dimexe
                                    W1.Range("B" & Row1 + 25).Value = style1.Dimexo
                                    W1.Range("B" & Row1 + 26).Value = style1.Dimfrac
                                    W1.Range("B" & Row1 + 27).Value = style1.Dimfxlen
                                    W1.Range("B" & Row1 + 28).Value = style1.DimfxlenOn
                                    W1.Range("B" & Row1 + 29).Value = style1.Dimgap
                                    W1.Range("B" & Row1 + 30).Value = style1.Dimjogang
                                    W1.Range("B" & Row1 + 31).Value = style1.Dimjust
                                    W1.Range("B" & Row1 + 32).Value = style1.Dimblk.ToString
                                    W1.Range("B" & Row1 + 33).Value = style1.Dimldrblk.ToString
                                    W1.Range("B" & Row1 + 34).Value = style1.Dimlfac
                                    W1.Range("B" & Row1 + 35).Value = style1.Dimlim
                                    W1.Range("B" & Row1 + 36).Value = style1.Dimlunit
                                    W1.Range("B" & Row1 + 37).Value = style1.Dimlwd
                                    W1.Range("B" & Row1 + 38).Value = style1.Dimlwe
                                    W1.Range("B" & Row1 + 39).Value = style1.Dimpost
                                    W1.Range("B" & Row1 + 40).Value = style1.Dimrnd
                                    W1.Range("B" & Row1 + 41).Value = style1.Dimsah
                                    W1.Range("B" & Row1 + 42).Value = style1.Dimscale
                                    W1.Range("B" & Row1 + 43).Value = style1.Dimsd1
                                    W1.Range("B" & Row1 + 44).Value = style1.Dimsd2
                                    W1.Range("B" & Row1 + 45).Value = style1.Dimse1
                                    W1.Range("B" & Row1 + 46).Value = style1.Dimse2
                                    W1.Range("B" & Row1 + 47).Value = style1.Dimsoxd
                                    W1.Range("B" & Row1 + 48).Value = style1.Dimtad
                                    W1.Range("B" & Row1 + 49).Value = style1.Dimtdec
                                    W1.Range("B" & Row1 + 50).Value = style1.Dimtfac
                                    W1.Range("B" & Row1 + 51).Value = style1.Dimtfill
                                    W1.Range("B" & Row1 + 52).Value = style1.Dimtfillclr.ColorIndex
                                    W1.Range("B" & Row1 + 53).Value = style1.Dimtih
                                    W1.Range("B" & Row1 + 54).Value = style1.Dimtix
                                    W1.Range("B" & Row1 + 55).Value = style1.Dimtm
                                    W1.Range("B" & Row1 + 56).Value = style1.Dimtmove
                                    W1.Range("B" & Row1 + 57).Value = style1.Dimtofl
                                    W1.Range("B" & Row1 + 58).Value = style1.Dimtoh
                                    W1.Range("B" & Row1 + 59).Value = style1.Dimtol
                                    W1.Range("B" & Row1 + 60).Value = style1.Dimtolj
                                    W1.Range("B" & Row1 + 61).Value = style1.Dimtp
                                    W1.Range("B" & Row1 + 62).Value = style1.Dimtsz
                                    W1.Range("B" & Row1 + 63).Value = style1.Dimtvp
                                    W1.Range("B" & Row1 + 64).Value = style1.Dimtxsty.ToString
                                    W1.Range("B" & Row1 + 65).Value = style1.Dimtxt
                                    W1.Range("B" & Row1 + 66).Value = style1.Dimtxtdirection
                                    W1.Range("B" & Row1 + 67).Value = style1.Dimtzin
                                    W1.Range("B" & Row1 + 68).Value = style1.Dimupt
                                    W1.Range("B" & Row1 + 69).Value = style1.Dimzin


                                    '.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                    '.Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                                    '.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)

                                    '.Dimdsep = ".c"

                                    '.Dimjogang = 0.785398163397448

                                    '.Dimblk = Arrowid_OPEN30
                                    '.Dimldrblk = Arrowid

                                    '.Dimlwd = LineWeight.ByBlock
                                    '.Dimlwe = LineWeight.ByBlock


                                    '.Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)

                                    '.Dimtxsty = Text_style_Text25.ObjectId



                                    Row1 = Row1 + 70
                                Case MsgBoxResult.Cancel
                                    Exit For
                            End Select
                        End If

                    Next
                End Using
            End Using

            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("water_keep")>
    Public Sub find_replace1()

        If isSECURE() = False Then Exit Sub

        Try
            'Application.SetSystemVariable("FILEDIA", 0)






            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument



            Dim Colectie_nume_layout As New Specialized.StringCollection

            Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Layoutdict As DBDictionary
                    Layoutdict = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)
                    For Each entry As DBDictionaryEntry In Layoutdict
                        If Not entry.Key.ToUpper = "MODEL" Then
                            Colectie_nume_layout.Add(entry.Key)
                        End If
                    Next

                    If Colectie_nume_layout.Count > 0 Then
                        Dim layMgr As LayoutManager = LayoutManager.Current
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)

                        Dim String_de_cautat As String = "ENTER/ EXIT WETLAND".ToUpper

                        For i = 0 To Colectie_nume_layout.Count - 1
                            Dim nume_layout As String = Colectie_nume_layout(i)
                            layMgr.CurrentLayout = nume_layout
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                            For Each ObjID In BTrecord
                                Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If TypeOf DBobject Is DBText Then
                                    Dim Text1 As DBText = DBobject
                                    Dim Continut As String = Text1.TextString.ToUpper

                                    If Continut.Contains(String_de_cautat) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat)
                                        Text1.UpgradeOpen()
                                        Text1.TextString = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat) - 1)
                                    End If


                                End If
                                If TypeOf DBobject Is MText Then
                                    Dim mText1 As MText = DBobject
                                    Dim Continut As String = mText1.Contents.ToUpper

                                    If Continut.Contains(String_de_cautat) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat)
                                        mText1.UpgradeOpen()
                                        mText1.Contents = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat) - 1)
                                    End If


                                End If
                            Next









                        Next
                    End If


                    Trans1.Commit()
                End Using
            End Using










            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            'SET_FILEDIA_TO_1()
            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("WL")>
    Public Sub find_replace2()

        If isSECURE() = False Then Exit Sub

        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument



            Dim Colectie_nume_layout As New Specialized.StringCollection

            Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim Layoutdict As DBDictionary
                    Layoutdict = Trans1.GetObject(ThisDrawing.Database.LayoutDictionaryId, OpenMode.ForRead)
                    For Each entry As DBDictionaryEntry In Layoutdict
                        If Not entry.Key.ToUpper = "MODEL" Then
                            Colectie_nume_layout.Add(entry.Key)
                        End If
                    Next

                    If Colectie_nume_layout.Count > 0 Then
                        Dim layMgr As LayoutManager = LayoutManager.Current
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)

                        Dim String_de_cautat1 As String = "ENTER/ EXIT WETLAND".ToUpper
                        Dim String_de_cautat2 As String = "ENTER/ EXIT WATERBODY".ToUpper
                        Dim String_de_cautat3 As String = "CENTERLINE OF STREAM".ToUpper
                        Dim String_de_cautat4 As String = "EDGE OF STREAM".ToUpper



                        For i = 0 To Colectie_nume_layout.Count - 1
                            Dim nume_layout As String = Colectie_nume_layout(i)
                            layMgr.CurrentLayout = nume_layout
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                            For Each ObjID In BTrecord
                                Dim DBobject As DBObject = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If TypeOf DBobject Is DBText Then
                                    Dim Text1 As DBText = DBobject
                                    Dim Continut As String = Text1.TextString.ToUpper

                                    If Continut.Contains(String_de_cautat1) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat1)
                                        Text1.UpgradeOpen()
                                        Text1.TextString = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat1) - 1)
                                    End If


                                    If Continut.Contains(String_de_cautat2) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat2)
                                        Text1.UpgradeOpen()
                                        Text1.TextString = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat2) - 1)
                                    End If


                                    If Continut.Contains(String_de_cautat3) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat3)
                                        Text1.UpgradeOpen()
                                        Text1.TextString = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat3) - 1)
                                    End If


                                    If Continut.Contains(String_de_cautat4) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat4)
                                        Text1.UpgradeOpen()
                                        Text1.TextString = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat4) - 1)
                                    End If


                                End If
                                If TypeOf DBobject Is MText Then
                                    Dim mText1 As MText = DBobject
                                    Dim Continut As String = mText1.Contents.ToUpper

                                    If Continut.Contains(String_de_cautat1) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat1)
                                        mText1.UpgradeOpen()
                                        mText1.Contents = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat1) - 1)
                                    End If


                                    If Continut.Contains(String_de_cautat2) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat2)
                                        mText1.UpgradeOpen()
                                        mText1.Contents = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat2) - 1)
                                    End If


                                    If Continut.Contains(String_de_cautat3) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat3)
                                        mText1.UpgradeOpen()
                                        mText1.Contents = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat3) - 1)
                                    End If


                                    If Continut.Contains(String_de_cautat4) = True Then
                                        Dim Pozitie1 As Integer = InStr(Continut, String_de_cautat4)
                                        mText1.UpgradeOpen()
                                        mText1.Contents = Strings.Left(Continut, Pozitie1 + Len(String_de_cautat4) - 1)
                                    End If



                                End If
                            Next









                        Next
                    End If


                    Trans1.Commit()
                End Using
            End Using










            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            'SET_FILEDIA_TO_1()
            MsgBox(ex.Message)
        End Try


    End Sub


    <CommandMethod("ENV_BAND_LEN")>
    Public Sub EDIT_ENV_BAND()

        If isSECURE() = False Then Exit Sub

        Try
            'Application.SetSystemVariable("FILEDIA", 0)

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
            Dim DeltaY As Double = 4.605

            Dim Colectie_nume_layout As New Specialized.StringCollection

            Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select mtexts:"

                    Object_Prompt.SingleOnly = False
                    Rezultat1 = Editor1.GetSelection(Object_Prompt)

                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        If IsNothing(Rezultat1) = False Then


                            For i = 0 To Rezultat1.Value.Count - 1

                                Dim DBobject As Entity = Trans1.GetObject(Rezultat1.Value.Item(i).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                If TypeOf DBobject Is MText Then
                                    Dim mText1 As MText = DBobject
                                    Dim Continut As String = mText1.Text.ToUpper
                                    Dim Continut1 As String = mText1.Contents.ToUpper
                                    Dim Continut2 As String
                                    If Len(Continut) > 46 Then
                                        Dim Pos As Integer
                                        If Continut1.Contains(vbCr) = False Then
                                            If Continut1.Contains("TRIBUTARY TO ") = True Then
                                                Pos = InStr(Continut1, "TRIBUTARY") + 13
                                            Else
                                                Pos = CInt(Len(Continut1) / 2 + (Len(Continut1) - Len(Continut)) / 2)
                                                For j = 1 To Pos
                                                    If Strings.Right(Left(Continut1, Pos - j), 1) = " " Then
                                                        Pos = Pos - j
                                                        Exit For
                                                    End If
                                                Next

                                            End If
                                            Continut2 = Left(Continut1, Pos - 1) & vbCrLf & "  " & Mid(Continut1, Pos)
                                            mText1.Contents = Continut2
                                            mText1.Location = New Point3d(mText1.Location.X, mText1.Location.Y - DeltaY, 0)
                                        End If



                                    End If

                                End If







                            Next
                        End If

                    End If

                    Trans1.Commit()
                End Using
            End Using










            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            'SET_FILEDIA_TO_1()
            MsgBox(ex.Message)
        End Try


    End Sub


    <CommandMethod("Point_at_station")>
    Public Sub point_at_station_us_style()

        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If
            Dim Poly2D As Polyline

            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d




            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1



                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim String1 As Autodesk.AutoCAD.EditorInput.PromptStringOptions
                String1 = New Autodesk.AutoCAD.EditorInput.PromptStringOptions(vbLf & "Specify station:")
                String1.AllowSpaces = True

                Dim Descriptia As Autodesk.AutoCAD.EditorInput.PromptResult = Editor1.GetString(String1)

                If Descriptia.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    Exit Sub
                End If

                Dim Ch_result As String = Descriptia.StringResult
                Ch_result = Replace(Ch_result, "+", "")
                Ch_result = Replace(Ch_result, " ", "")
                If IsNumeric(Ch_result) = False Then
                    MsgBox("Station is not specified correctly")
                    Exit Sub

                End If
                Dim Chainage As Double = CDbl(Ch_result)

                If Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If
                If IsNothing(Poly2D) = False Then
                    Point_on_poly = Poly2D.GetPointAtDist(Chainage)
                End If
                If IsNothing(Poly3D) = False Then
                    Point_on_poly = Poly3D.GetPointAtDist(Chainage)
                End If

                Dim Chainage_string As String = Get_chainage_feet_from_double(Chainage, 2)


                If Chainage_string = "-0+00" Then Chainage_string = "0+00"

                If IsNothing(Point_on_poly) = False Then
                    Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 50, 2.5, 20, 50, 100)
                End If

                Trans1.Commit()
                zoom_to_Point(Point_on_poly)

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("station_at_point")>
    Public Sub station_at_point_us_style()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly2D As Polyline
            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1


                            Dim Point_zero As New Point3d

                            Point_zero = Poly2D.GetClosestPointTo(Poly2D.StartPoint, Vector3d.ZAxis, False)



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Dim Index_Poly As Integer = 0
                            Poly2D = New Polyline
                            For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In Poly3D
                                Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                        (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                Dim x1 As Double = v3d.Position.X
                                Dim y1 As Double = v3d.Position.Y
                                Dim z1 As Double = v3d.Position.Z
                                Poly2D.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                Index_Poly = Index_Poly + 1
                            Next



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the same polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly2D) = False Then
                    Point_on_poly = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly2D.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D) = False Then
                    Dim Point_on_poly2D As New Point3d
                    Point_on_poly2D = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Dim param1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2D)
                    Distanta_pana_la_xing = Poly3D.GetDistanceAtParameter(param1)
                    Point_on_poly = Poly3D.GetPointAtParameter(param1)
                End If


                Dim Chainage As Double = Distanta_pana_la_xing - Dist_from_start_for_zero




                If Dist_from_start_for_zero + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If




                Dim Chainage_string As String = Get_chainage_feet_from_double(Chainage, 2)
                If Chainage_string = "-0+00" Then Chainage_string = "0+00"

                Dim Mleader1 As New MLeader

                If IsNothing(Point_on_poly) = False Then
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 3, 2.5, 3, 0, 5)
                End If


                Trans1.Commit()
                ''zoom_to_Point(Point_on_poly)
                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("sap_iso")>
    Public Sub station_at_point_isometric_style()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly2D As Polyline
            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_for_zero As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1


                            Dim Point_zero As New Point3d

                            Point_zero = Poly2D.GetClosestPointTo(Poly2D.StartPoint, Vector3d.ZAxis, False)



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Dim Index_Poly As Integer = 0
                            Poly2D = New Polyline
                            For Each vId As Autodesk.AutoCAD.DatabaseServices.ObjectId In Poly3D
                                Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = DirectCast(Trans1.GetObject _
                                        (vId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d)

                                Dim x1 As Double = v3d.Position.X
                                Dim y1 As Double = v3d.Position.Y
                                Dim z1 As Double = v3d.Position.Z
                                Poly2D.AddVertexAt(Index_Poly, New Point2d(x1, y1), 0, 0, 0)
                                Index_Poly = Index_Poly + 1
                            Next



                            Dist_from_start_for_zero = 0

                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the same polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly2D) = False Then
                    Point_on_poly = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly2D.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D) = False Then
                    Dim Point_on_poly2D As New Point3d
                    Point_on_poly2D = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Dim param1 As Double = Poly2D.GetParameterAtPoint(Point_on_poly2D)
                    Distanta_pana_la_xing = Poly3D.GetDistanceAtParameter(param1)
                    Point_on_poly = Poly3D.GetPointAtParameter(param1)
                End If


                Dim Chainage As Double = Distanta_pana_la_xing - Dist_from_start_for_zero




                If Dist_from_start_for_zero + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If




                Dim Chainage_string As String = Get_chainage_feet_from_double(Chainage, 0)
                If Chainage_string = "-0+00" Then Chainage_string = "0+00"



                If IsNothing(Point_on_poly) = False Then
                    Dim Mtext1 As New MText
                    Mtext1.TextHeight = 40
                    Mtext1.Rotation = 330 * PI / 180
                    Mtext1.Attachment = AttachmentPoint.TopLeft
                    Mtext1.Location = Point_on_poly
                    Mtext1.Contents = "\pxql;{\Q-30;" & Chainage_string & "}"
                    BTrecord.AppendEntity(Mtext1)
                    Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                End If


                Trans1.Commit()

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub


    <CommandMethod("ppl_DEFL_USA", CommandFlags.UsePickSet)>
    Public Sub DEFLECTION_FROM_2D_USA()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 2D polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline Then
                            Dim Poly2D As Polyline = Ent1

                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "STATION"
                            W1.Range("C" & 1).Value = "X"
                            W1.Range("D" & 1).Value = "Y"
                            W1.Range("E" & 1).Value = "Z"
                            W1.Range("F" & 1).Value = "DEFLECTION"
                            W1.Range("G" & 1).Value = "DEFLECTION DD"
                            W1.Range("H" & 1).Value = "PI"
                            For j = 0 To Poly2D.NumberOfVertices - 1



                                If j > 0 And j < Poly2D.NumberOfVertices - 1 Then


                                    Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D.NumberOfVertices - 1 Then
                                                vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                    Dim Mleader1 As New MLeader
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



                                    Dim Degree As Integer = Floor(angle1)
                                    Dim Min1 As Integer = Floor((angle1 - Floor(angle1)) * 60)
                                    Dim Sec1 As Integer = Round(((angle1 - Degree) * 60 - Min1) * 60, 0)
                                    If Sec1 = 60 Then
                                        Sec1 = 0
                                        Min1 = Min1 + 1
                                    End If
                                    If Min1 = 60 Then
                                        Degree = Degree + 1
                                        Min1 = 0
                                    End If

                                    Dim Minute As String = Min1.ToString
                                    If Len(Minute) = 1 Then Minute = "0" & Minute
                                    Dim Second As String = Sec1.ToString
                                    If Len(Second) = 1 Then Second = "0" & Second

                                    Dim AngleDMS As String = Degree.ToString & "°" & Minute & "'" & Second & Chr(34)



                                    Dim continut_mleader As String = AngleDMS & LT_RT

                                    If angle1 >= 0.5 Then

                                        W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                        W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                        W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                        W1.Range("E" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)
                                        W1.Range("F" & RowXL).Value = continut_mleader
                                        W1.Range("G" & RowXL).Value = Round(angle1, 5)
                                        W1.Range("H" & RowXL).Value = "P.I. ~ " & continut_mleader

                                        Creaza_Mleader_nou_fara_UCS_transform(Poly2D.GetPointAtParameter(j),
                                                                              Round(Poly2D.GetDistanceAtParameter(j), 0) & vbCrLf &
                                                                              "X = " & Round(Poly2D.GetPointAtParameter(j).X, 3) & vbCrLf &
                                                                              "Y = " & Round(Poly2D.GetPointAtParameter(j).Y, 3) & vbCrLf &
                                                                              "Z = " & Round(Poly2D.GetPointAtParameter(j).Z, 3) & vbCrLf &
                                                                              continut_mleader, 2.5, 2.5, 2.5, 5, 10)


                                        RowXL = RowXL + 1

                                    End If
                                Else
                                    W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                    W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                    W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                    W1.Range("E" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)

                                    RowXL = RowXL + 1
                                End If
                            Next
                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    <CommandMethod("ppl_DEFL_USA0", CommandFlags.UsePickSet)>
    Public Sub DEFLECTION_FROM_2D_all()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select 2D polyline:"

                Object_Prompt.SingleOnly = True
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)


                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline Then
                            Dim Poly2D As Polyline = Ent1

                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "STATION"
                            W1.Range("C" & 1).Value = "X"
                            W1.Range("D" & 1).Value = "Y"
                            W1.Range("E" & 1).Value = "Z"
                            W1.Range("F" & 1).Value = "DEFLECTION"
                            W1.Range("G" & 1).Value = "DEFLECTION DD"
                            W1.Range("H" & 1).Value = "PI"
                            For j = 0 To Poly2D.NumberOfVertices - 1



                                If j > 0 And j < Poly2D.NumberOfVertices - 1 Then


                                    Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D.NumberOfVertices - 1 Then
                                                vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                    Dim Mleader1 As New MLeader
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



                                    Dim Degree As Integer = Floor(angle1)
                                    Dim Min1 As Integer = Floor((angle1 - Floor(angle1)) * 60)
                                    Dim Sec1 As Integer = Round(((angle1 - Degree) * 60 - Min1) * 60, 0)
                                    If Sec1 = 60 Then
                                        Sec1 = 0
                                        Min1 = Min1 + 1
                                    End If
                                    If Min1 = 60 Then
                                        Degree = Degree + 1
                                        Min1 = 0
                                    End If

                                    Dim Minute As String = Min1.ToString
                                    If Len(Minute) = 1 Then Minute = "0" & Minute
                                    Dim Second As String = Sec1.ToString
                                    If Len(Second) = 1 Then Second = "0" & Second

                                    Dim AngleDMS As String = Degree.ToString & "°" & Minute & "'" & Second & Chr(34)



                                    Dim continut_mleader As String = AngleDMS & LT_RT



                                    W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                    W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                    W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                    W1.Range("E" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)
                                    W1.Range("F" & RowXL).Value = continut_mleader
                                    W1.Range("G" & RowXL).Value = Round(angle1, 5)
                                    W1.Range("H" & RowXL).Value = "P.I. ~ " & continut_mleader

                                    Creaza_Mleader_nou_fara_UCS_transform(Poly2D.GetPointAtParameter(j),
                                                                          Round(Poly2D.GetDistanceAtParameter(j), 0) & vbCrLf &
                                                                          "X = " & Round(Poly2D.GetPointAtParameter(j).X, 3) & vbCrLf &
                                                                          "Y = " & Round(Poly2D.GetPointAtParameter(j).Y, 3) & vbCrLf &
                                                                          "Z = " & Round(Poly2D.GetPointAtParameter(j).Z, 3) & vbCrLf &
                                                                          continut_mleader, 2.5, 2.5, 2.5, 5, 10)


                                    RowXL = RowXL + 1


                                Else
                                    W1.Range("A" & RowXL).Value = Poly2D.GetDistanceAtParameter(j)
                                    W1.Range("C" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).X, 3)
                                    W1.Range("D" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Y, 3)
                                    W1.Range("E" & RowXL).Value = Round(Poly2D.GetPointAtParameter(j).Z, 3)

                                    RowXL = RowXL + 1
                                End If
                            Next
                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("PPL_ADDRR", CommandFlags.UsePickSet)>
    Public Sub add_rr_to_chainages()

        If isSECURE() = False Then Exit Sub

        Try
            'Application.SetSystemVariable("FILEDIA", 0)






            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument


            Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Rezultat1 = ThisDrawing.Editor.SelectImplied

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    Else
                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select attribute blocks, Text, Mtext objects:"

                        Object_Prompt.SingleOnly = True
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Object_Prompt)
                    End If

                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If


                    Dim String_de_cautat1 As String = "+"
                    Dim String_de_cautat2 As String = "."
                    Dim String_de_adaugat As String = "RR"

                    For i = 0 To Rezultat1.Value.Count - 1
                        Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value.Item(i).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If TypeOf Ent1 Is DBText Then
                            Dim Text1 As DBText = Ent1
                            Dim Continut As String = Text1.TextString.ToUpper

                            If Continut.Contains(String_de_cautat1) = True And Continut.Contains(String_de_cautat2) = True And Continut.Contains(String_de_adaugat) = False Then

                                Text1.UpgradeOpen()
                                Text1.TextString = Text1.TextString & String_de_adaugat
                            End If


                        End If
                        If TypeOf Ent1 Is MText Then
                            Dim mText1 As MText = Ent1
                            Dim Continut As String = mText1.Contents.ToUpper

                            If Continut.Contains(String_de_cautat1) = True And Continut.Contains(String_de_cautat2) = True And Continut.Contains(String_de_adaugat) = False Then

                                mText1.UpgradeOpen()
                                mText1.Contents = Replace(mText1.Contents, mText1.Text, mText1.Text & String_de_adaugat)
                            End If


                        End If
                        If TypeOf Ent1 Is BlockReference Then
                            Dim Block1 As BlockReference = Ent1

                            If Block1.AttributeCollection.Count > 0 Then
                                Block1.UpgradeOpen()
                                For Each Id1 As ObjectId In Block1.AttributeCollection
                                    Dim Atr1 As AttributeReference
                                    Atr1 = TryCast(Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), AttributeReference)
                                    If IsNothing(Atr1) = False Then
                                        Dim Tag1 As String = Atr1.Tag
                                        Dim Continut As String
                                        If Atr1.IsMTextAttribute = True Then
                                            Continut = Atr1.MTextAttribute.Contents
                                        Else
                                            Continut = Atr1.TextString
                                        End If
                                        If Continut.Contains(String_de_cautat1) = True And Continut.Contains(String_de_cautat2) = True And Continut.Contains(String_de_adaugat) = False Then
                                            If Atr1.IsMTextAttribute = True Then
                                                Atr1.MTextAttribute.Contents = Replace(Atr1.MTextAttribute.Contents, Continut, Continut & String_de_adaugat)
                                            Else
                                                Atr1.TextString = Atr1.TextString & String_de_adaugat
                                            End If
                                        End If

                                    End If
                                Next


                            End If




                        End If




                    Next











                    Trans1.Commit()
                End Using
            End Using










            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            'SET_FILEDIA_TO_1()
            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("PPL_MULT_HATCH", CommandFlags.UsePickSet)>
    Public Sub multiple_hatch()
        If isSECURE() = False Then Exit Sub

        Try
            'Application.SetSystemVariable("FILEDIA", 0)






            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument


            Using Lock1 As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                    Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                    Rezultat1 = ThisDrawing.Editor.SelectImplied

                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                    Else
                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select polylines:"

                        Object_Prompt.SingleOnly = True
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Object_Prompt)
                    End If

                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select hatch:"

                    Object_Prompt2.SingleOnly = True
                    Rezultat2 = ThisDrawing.Editor.GetSelection(Object_Prompt2)


                    If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If
                    Dim Ent2 As Entity = Trans1.GetObject(Rezultat2.Value.Item(0).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim Hatch2 As Hatch

                    If TypeOf Ent2 Is Hatch Then
                        Hatch2 = Ent2
                    End If
                    If IsNothing(Hatch2) = True Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    For i = 0 To Rezultat1.Value.Count - 1
                        Dim Ent1 As Entity = Trans1.GetObject(Rezultat1.Value.Item(i).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If TypeOf Ent1 Is Polyline Then
                            Dim Poly1 As Polyline = Ent1
                            If Poly1.Closed = False Then
                                Poly1.UpgradeOpen()
                                Poly1.Closed = True
                            End If

                            Dim Hatch1 As New Hatch
                            BTrecord.AppendEntity(Hatch1)
                            Trans1.AddNewlyCreatedDBObject(Hatch1, True)
                            Hatch1.SetHatchPattern(HatchPatternType.PreDefined, Hatch2.PatternName)

                            Hatch1.PatternScale = Hatch2.PatternScale
                            Hatch1.PatternAngle = Hatch2.PatternAngle
                            Hatch1.PatternSpace = Hatch2.PatternSpace

                            Dim oBJiD_COL_H As New ObjectIdCollection
                            oBJiD_COL_H.Add(Ent1.ObjectId)
                            Hatch1.AppendLoop(HatchLoopTypes.External, oBJiD_COL_H)
                            Hatch1.Layer = Hatch2.Layer
                            Hatch1.EvaluateHatch(True)


                        End If










                    Next











                    Trans1.Commit()
                End Using
            End Using










            'Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Catch ex As Exception
            'SET_FILEDIA_TO_1()
            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("PPL_AREA")>
    Public Sub area_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Area_Form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Area_Form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ' <CommandMethod("ATP", CommandFlags.UsePickSet)>
    Public Sub Align_text_to_points()
        If isSECURE() = False Then Exit Sub
        Dim NEW_OSnap, Old_OSnap As Integer
        Old_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

        NEW_OSnap = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select text:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify first point:")

                    Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify second point:")

                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)

                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                        Exit Sub
                    End If


                    Dim X1, X2 As Double
                    Dim Y1, Y2 As Double


                    X1 = Point1.Value.X
                    Y1 = Point1.Value.Y

                    PP2.BasePoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y1, 0)
                    PP2.UseBasePoint = True
                    PP2.AllowNone = False
                    Point2 = Editor1.GetPoint(PP2)

                    If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                        Exit Sub
                    End If


                    X2 = Point2.Value.X
                    Y2 = Point2.Value.Y

                    Dim Rotatie2 As Double

                    Rotatie2 = GET_Bearing_rad(X1, Y1, X2, Y2)

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Mtext1 As MText
                        Dim Text1 As DBText
                        Dim Block1 As BlockReference
                        Dim Mleader1 As Autodesk.AutoCAD.DatabaseServices.MLeader

                        For i = 1 To Rezultat1.Value.Count

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i - 1)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                Ent1.UpgradeOpen()
                                Mtext1 = Ent1
                                Mtext1.Rotation = Rotatie2

                            End If

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                Ent1.UpgradeOpen()
                                Text1 = Ent1
                                Text1.Rotation = Rotatie2

                            End If

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                Ent1.UpgradeOpen()
                                Block1 = Ent1
                                'Block1.Rotation = Rotatie2

                                Dim x0 As Double = Block1.Position.X
                                Dim y0 As Double = Block1.Position.Y



                                Dim Rotation1 As Double = Block1.Rotation
                                Dim Rotation2 As Double = GET_Bearing_rad(X1, Y1, X2, Y2)
                                Dim Rotation_result As Double = Rotation2 - Rotation1



                                Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                                Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                                Ent1.TransformBy(Matrix3d.Rotation(Rotation_result, CurentUCS.Zaxis, New Point3d(x0, y0, 0)))


                            End If

                            If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Wipeout Then
                                Ent1.UpgradeOpen()
                                Dim Wipe1 As Wipeout = Ent1

                                If Wipe1.HasFrame = False Then
                                    ' Wipe1.HasFrame = True
                                Else
                                    'Wipe1.HasFrame = False
                                End If



                            End If

                        Next

                        Editor1.Regen()
                        Trans1.Commit()

                        Mtext1 = Nothing
                        Text1 = Nothing
                        Block1 = Nothing
                        Mleader1 = Nothing

                    End Using
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                Else
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub
    <CommandMethod("Ao2P", CommandFlags.UsePickSet)>
    Public Sub Align_object_to_points()
        If isSECURE() = False Then Exit Sub
        Dim NEW_OSnap, Old_OSnap, Base_Osnap As Integer
        Old_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE")

        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then
                    Base_Osnap = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Intersection + Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.End

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Base_Osnap)

                    Dim x0, y0 As Double
                    Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Specify rotation base:")
                    PP0.AllowNone = False
                    Point0 = Editor1.GetPoint(PP0)

                    If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                        Exit Sub
                    End If
                    x0 = Point0.Value.X
                    y0 = Point0.Value.Y

                    NEW_OSnap = Autodesk.AutoCAD.EditorInput.ObjectSnapMasks.Near
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap)

                    Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Specify first FROM point:")

                    Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Specify second FROM point:")

                    PP1.AllowNone = False
                    Point1 = Editor1.GetPoint(PP1)

                    If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                        Exit Sub
                    End If


                    Dim X1, X2, Xx1, Xx2 As Double
                    Dim Y1, Y2, Yy1, Yy2 As Double


                    X1 = Point1.Value.X
                    Y1 = Point1.Value.Y

                    PP2.BasePoint = New Autodesk.AutoCAD.Geometry.Point3d(X1, Y1, 0)
                    PP2.UseBasePoint = True
                    PP2.AllowNone = False
                    Point2 = Editor1.GetPoint(PP2)

                    If Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                        Exit Sub
                    End If


                    X2 = Point2.Value.X
                    Y2 = Point2.Value.Y


                    Dim PointP1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                    Dim PointP2 As Autodesk.AutoCAD.EditorInput.PromptPointResult

                    Dim PPp1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify first TO point:")

                    Dim PPp2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions("Specify second TO point:")

                    PPp1.AllowNone = False
                    PointP1 = Editor1.GetPoint(PPp1)

                    If PointP1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                        Exit Sub
                    End If


                    Xx1 = PointP1.Value.X
                    Yy1 = PointP1.Value.Y

                    PPp2.BasePoint = New Autodesk.AutoCAD.Geometry.Point3d(Xx1, Yy1, 0)
                    PPp2.UseBasePoint = True
                    PPp2.AllowNone = False
                    PointP2 = Editor1.GetPoint(PPp2)

                    If PointP2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                        Exit Sub
                    End If


                    Xx2 = PointP2.Value.X
                    Yy2 = PointP2.Value.Y


                    Dim Rotation1 As Double = GET_Bearing_rad(X1, Y1, X2, Y2)
                    Dim Rotation2 As Double = GET_Bearing_rad(Xx1, Yy1, Xx2, Yy2)
                    Dim Rotation_result As Double = Rotation2 - Rotation1


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction


                        For i = 1 To Rezultat1.Value.Count

                            Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                            Obj1 = Rezultat1.Value.Item(i - 1)
                            Dim Ent1 As Entity
                            Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                            Ent1.UpgradeOpen()

                            Dim CurentUCSmatrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
                            Dim CurentUCS As CoordinateSystem3d = CurentUCSmatrix.CoordinateSystem3d

                            Ent1.TransformBy(Matrix3d.Rotation(Rotation_result, CurentUCS.Zaxis, New Point3d(x0, y0, 0)))

                        Next

                        Editor1.Regen()
                        Trans1.Commit()



                    End Using
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                Else
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", Old_OSnap)
            'Exit Sub
            MsgBox(ex.Message)
        End Try


    End Sub

    <CommandMethod("Blocks2mtexts", CommandFlags.UsePickSet)>
    Public Sub Transfer_block_attributes_to_a_Mtext()
        If isSECURE() = False Then Exit Sub
        Try

            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select the blocks containing {Line1,LINE2} and {CHAINAGE}:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Lock1 As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument

                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Creaza_layer("NO PLOT", 40, "", False)

                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForWrite)
                                Dim BlockTable_data1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                                BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)



                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                    Using BlockReference1 As BlockReference = Ent1


                                        Dim Colectie_atribute As AttributeCollection = BlockReference1.AttributeCollection
                                        If Colectie_atribute.Count > 0 Then
                                            Dim Chainage As String = ""
                                            Dim Description1 As String
                                            Dim Description2 As String

                                            For Each id1 As ObjectId In Colectie_atribute
                                                Dim Atribut_ref As AttributeReference = Trans1.GetObject(id1, OpenMode.ForRead)
                                                If Atribut_ref.Tag = "CHAINAGE" Then
                                                    Chainage = Atribut_ref.TextString
                                                End If
                                                If Atribut_ref.Tag = "FIRST_LINE" Then
                                                    Description1 = Atribut_ref.TextString.ToUpper
                                                End If
                                                If Atribut_ref.Tag = "SECOND_LINE" Then
                                                    Description2 = Atribut_ref.TextString.ToUpper
                                                End If
                                            Next

                                            Dim Description As String
                                            If Description1 = "" Then
                                                If Not Description2 = "" Then
                                                    Description = Description2
                                                End If

                                            End If
                                            If Not Description1 = "" Then
                                                If Description2 = "" Then
                                                    Description = Description1
                                                Else
                                                    Description = Description1 & " " & Description2
                                                End If

                                            End If


                                            If Not Chainage = "" Then
                                                Dim String1 As String = Description & " " & Chainage
                                                Dim Mtext1 As New MText
                                                Mtext1.Contents = String1
                                                Mtext1.Layer = "NO PLOT"
                                                Mtext1.TextHeight = 1
                                                Mtext1.Location = BlockReference1.Position
                                                BTrecord.AppendEntity(Mtext1)
                                                Trans1.AddNewlyCreatedDBObject(Mtext1, True)




                                            End If

                                        End If
                                    End Using
                                End If

                            Next


                            Trans1.Commit()
                        End Using


                    End Using 'asta e de la Using Lock1
                Else
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            'Exit Sub
            MsgBox(ex.Message)
        End Try
    End Sub



    <CommandMethod("ppl_layer2XL")>
    Public Sub LAYERS_TO_EXCEL()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try





            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim Layer_table As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                Dim RowXL As Integer = 2
                W1 = get_new_worksheet_from_Excel()
                W1.Range("A" & 1).Value = "NAME"
                W1.Range("C" & 1).Value = "ON-OFF"
                W1.Range("D" & 1).Value = "THAW-FREEZE"
                W1.Range("B" & 1).Value = "COLOR"
                W1.Range("E" & 1).Value = "PLOTABLE"
                W1.Range("F" & 1).Value = "LINETYPE"


                For Each ID1 As ObjectId In Layer_table
                    Dim LayerTRec1 As LayerTableRecord = Trans1.GetObject(ID1, OpenMode.ForRead)
                    W1.Range("A" & RowXL).Value = LayerTRec1.Name
                    W1.Range("B" & RowXL).Value = LayerTRec1.Color.ColorIndex
                    If LayerTRec1.IsOff = True Then
                        W1.Range("C" & RowXL).Value = "OFF"
                    Else
                        W1.Range("C" & RowXL).Value = "ON"
                    End If
                    If LayerTRec1.IsFrozen = True Then
                        W1.Range("D" & RowXL).Value = "FROZEN"
                    Else
                        W1.Range("D" & RowXL).Value = "THAW"
                    End If
                    If LayerTRec1.IsPlottable = True Then
                        W1.Range("E" & RowXL).Value = "PLOTABLE"
                    Else
                        W1.Range("E" & RowXL).Value = "NON PLOTABLE"
                    End If
                    Dim Ltype As LinetypeTableRecord = Trans1.GetObject(LayerTRec1.LinetypeObjectId, OpenMode.ForRead)
                    W1.Range("F" & RowXL).Value = Ltype.Name

                    RowXL = RowXL + 1
                Next



                Trans1.Commit()
            End Using


            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    <CommandMethod("PPL_BUILD_PI")>
    Public Sub BUILD_PI_on_3D()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Editor1.SetImpliedSelection(Empty_array)

        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select 3D polyline:"

            Object_Prompt.SingleOnly = True
            Rezultat1 = Editor1.GetSelection(Object_Prompt)

            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Dim Param1 As Integer
                    Dim Param2 As Integer
                    Dim Poly1 As New Polyline
                    Dim PI1 As New Point3d


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)



                        Dim Ent1 As Entity
                        Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)


                        If TypeOf Ent1 Is Polyline3d Then

                            Dim Poly2D As New Polyline
                            Dim Poly3D As Polyline3d = Ent1

                            Dim Index2d As Double = 0
                            For Each ObjId As ObjectId In Poly3D
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d = Index2d + 1
                            Next
                            Poly2D.Elevation = 0

                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick 1st point:")
                            Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            PP1.AllowNone = False
                            Point1 = Editor1.GetPoint(PP1)
                            If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If


                            Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick 2nd point:")
                            Dim Point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            PP2.UseBasePoint = True
                            PP2.BasePoint = Point1.Value
                            PP2.AllowNone = False
                            Point2 = Editor1.GetPoint(PP2)
                            If Not Point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If

                            Dim Point_on_Poly2d1 As Point3d
                            Point_on_Poly2d1 = Poly2D.GetClosestPointTo(Point1.Value, Vector3d.ZAxis, False)
                            Param1 = Poly2D.GetParameterAtPoint(Point_on_Poly2d1)

                            Dim Point_on_Poly2d2 As Point3d
                            Point_on_Poly2d2 = Poly2D.GetClosestPointTo(Point2.Value, Vector3d.ZAxis, False)
                            Param2 = Poly2D.GetParameterAtPoint(Point_on_Poly2d2)


                            If Param1 > Param2 Then
                                Dim T As Integer = Param1
                                Param1 = Param2
                                Param2 = T


                            End If

                            If Param2 - Param1 < 1 Then
                                MsgBox("Parameters too close")
                                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If



                            Poly1.AddVertexAt(0, Poly2D.GetPoint2dAt(Param1 - 1), 0, 0, 0)
                            Poly1.AddVertexAt(1, Poly2D.GetPoint2dAt(Param1), 0, 0, 0)
                            Poly1.Elevation = 0

                            Dim Poly2 As New Polyline
                            Poly2.AddVertexAt(0, Poly2D.GetPoint2dAt(Param2 + 1), 0, 0, 0)
                            Poly2.AddVertexAt(1, Poly2D.GetPoint2dAt(Param2), 0, 0, 0)
                            Poly2.Elevation = 0

                            Dim Col_pt As New Point3dCollection
                            Poly1.IntersectWith(Poly2, Intersect.ExtendBoth, Col_pt, IntPtr.Zero, IntPtr.Zero)
                            If Col_pt.Count = 1 Then
                                PI1 = Col_pt(0)
                                Poly1.AddVertexAt(2, New Point2d(PI1.X, PI1.Y), 0, 0, 0)
                                Poly1.AddVertexAt(3, Poly2D.GetPoint2dAt(Param2), 0, 0, 0)
                                Poly1.AddVertexAt(4, Poly2D.GetPoint2dAt(Param2 + 1), 0, 0, 0)

                                Poly1.ColorIndex = 1
                                BTrecord.AppendEntity(Poly1)
                                Trans1.AddNewlyCreatedDBObject(Poly1, True)
                            End If

                        Else
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If

                        Trans1.Commit()
                    End Using

                    Editor1.Regen()

                    If MsgBox("Modify 3d Poly?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            Dim Ent1 As Entity
                            Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForWrite)


                            If TypeOf Ent1 Is Polyline3d Then
                                Dim Poly3D As Polyline3d = Ent1
                                Dim Poly_int As Polyline = Trans1.GetObject(Poly1.ObjectId, OpenMode.ForWrite)

                                Dim New_poly3d As New Polyline3d
                                New_poly3d.SetDatabaseDefaults()
                                New_poly3d.Layer = Poly3D.Layer


                                BTrecord.AppendEntity(New_poly3d)
                                Trans1.AddNewlyCreatedDBObject(New_poly3d, True)


                                For Each ObjId As ObjectId In Poly3D
                                    Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                                    Dim Param_on3D As Double = Poly3D.GetParameterAtPoint(vertex1.Position)
                                    If Param_on3D <= Param1 Or Param_on3D >= Param2 Then


                                        Dim Vertex_new As New PolylineVertex3d(vertex1.Position)
                                        New_poly3d.AppendVertex(Vertex_new)
                                        Trans1.AddNewlyCreatedDBObject(Vertex_new, True)

                                        If Param_on3D = Param1 Then
                                            Dim Pt1 As New Point3d
                                            Pt1 = Poly3D.GetPointAtParameter(Param1)

                                            Dim Vertex_new1 As New PolylineVertex3d(New Point3d(PI1.X, PI1.Y, Pt1.Z))
                                            New_poly3d.AppendVertex(Vertex_new1)
                                            Trans1.AddNewlyCreatedDBObject(Vertex_new1, True)


                                        End If

                                    End If

                                Next

                                Poly3D.Erase()
                                Poly_int.Erase()

                                Trans1.Commit()
                            End If

                        End Using

                    End If

                Else

                    Exit Sub
                End If
            End If

            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub



    <CommandMethod("PPL_stationing")>
    Public Sub creaza_stationing_along_2d_polyline()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Editor1.SetImpliedSelection(Empty_array)

        Dim Line_length As Double = 10
        Dim Dist_CL_to_text As Double = 10.36
        Dim Mtext_height As Double = 8

        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select the stationing polyline:"

            Object_Prompt.SingleOnly = True
            Rezultat1 = Editor1.GetSelection(Object_Prompt)

            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)


                        Dim Ent1 As Entity
                        Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)



                        If TypeOf Ent1 Is Polyline Then
                            Dim Poly2D As Polyline = Ent1
                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_zero As New Point3d
                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select 0+000:")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)
                            If Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Point_zero = Poly2D.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                                Dim Point_dir As New Point3d
                                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select last station")
                                PP1.UseBasePoint = True
                                PP1.BasePoint = Point0.Value
                                PP1.AllowNone = True
                                Point1 = Editor1.GetPoint(PP1)
                                If Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                    Point_dir = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)

                                    Dim Sta1 As Double = Poly2D.GetDistAtPoint(Point_zero)
                                    Dim Sta2 As Double = Poly2D.GetDistAtPoint(Point_dir)
                                    If Round(Sta1, 0) = Round(Sta2, 0) Then
                                        MsgBox("you picked the same point")
                                        Exit Sub
                                    End If
                                    If Sta2 > Sta1 Then
                                        Dim number_of_labels As Integer = Ceiling((Sta2 - Sta1) / 100)
                                        For i = 0 To number_of_labels
                                            Dim Pt_on_Poly As Point3d
                                            If Poly2D.Length >= Sta1 + i * 100 Then


                                                Pt_on_Poly = Poly2D.GetPointAtDist(Sta1 + i * 100)
                                                Dim Param_pt As Double = Poly2D.GetParameterAtPoint(Pt_on_Poly)

                                                Dim Pt_before As Point3d
                                                Dim Pt_after As Point3d
                                                If Param_pt = Round(Param_pt, 0) And Param_pt >= 1 And Param_pt <= Poly2D.NumberOfVertices - 2 Then
                                                    Pt_before = Poly2D.GetPointAtParameter(Param_pt)
                                                    Pt_after = Poly2D.GetPointAtParameter(Param_pt + 1)
                                                ElseIf Param_pt = 0 Then
                                                    Pt_before = Poly2D.GetPointAtParameter(0)
                                                    Pt_after = Poly2D.GetPointAtParameter(1)
                                                ElseIf Param_pt = Poly2D.NumberOfVertices - 1 Then
                                                    Pt_before = Poly2D.GetPointAtParameter(Param_pt - 1)
                                                    Pt_after = Poly2D.GetPointAtParameter(Param_pt)
                                                Else
                                                    Pt_before = Poly2D.GetPointAtParameter(Floor(Param_pt))
                                                    Pt_after = Poly2D.GetPointAtParameter(Ceiling(Param_pt))
                                                End If

                                                Dim Mtext_rotation As Double
                                                Mtext_rotation = GET_Bearing_rad(Pt_before.X, Pt_before.Y, Pt_after.X, Pt_after.Y)

                                                Dim Line1 As New Line(Pt_on_Poly, Pt_after)
                                                Dim Circle1 As New Circle(Pt_on_Poly, Vector3d.ZAxis, Line_length)
                                                Dim Col_int1 As New Point3dCollection
                                                Line1.IntersectWith(Circle1, Intersect.OnBothOperands, Col_int1, IntPtr.Zero, IntPtr.Zero)
                                                Dim Line2 As New Line
                                                If Col_int1.Count > 0 Then
                                                    Line2 = New Line(Pt_on_Poly, Col_int1(0))
                                                    Line2.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt_on_Poly))
                                                    Line2.TransformBy(Matrix3d.Displacement(Line2.GetPointAtDist(Line2.Length / 2).GetVectorTo(Pt_on_Poly)))
                                                    BTrecord.AppendEntity(Line2)
                                                    Trans1.AddNewlyCreatedDBObject(Line2, True)
                                                Else
                                                    Line1.IntersectWith(Circle1, Intersect.ExtendThis, Col_int1, IntPtr.Zero, IntPtr.Zero)
                                                    If Col_int1.Count > 0 Then
                                                        Line2 = New Line(Pt_on_Poly, Col_int1(0))
                                                        Line2.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt_on_Poly))
                                                        Line2.TransformBy(Matrix3d.Displacement(Line2.GetPointAtDist(Line2.Length / 2).GetVectorTo(Pt_on_Poly)))
                                                        BTrecord.AppendEntity(Line2)
                                                        Trans1.AddNewlyCreatedDBObject(Line2, True)
                                                    End If
                                                End If
                                                If IsNothing(Line2) = False Then
                                                    Dim Circle2 As New Circle(Pt_on_Poly, Vector3d.ZAxis, Dist_CL_to_text)
                                                    Dim Col_int2 As New Point3dCollection
                                                    Line2.IntersectWith(Circle2, Intersect.ExtendThis, Col_int2, IntPtr.Zero, IntPtr.Zero)
                                                    If Col_int2.Count > 0 Then
                                                        Dim Pt_ins As New Point3d
                                                        Pt_ins = Col_int2(0)
                                                        Dim Mtext1 As New MText
                                                        Mtext1.Location = Pt_ins
                                                        Mtext1.Rotation = Mtext_rotation
                                                        Mtext1.TextHeight = Mtext_height
                                                        Mtext1.Contents = Get_chainage_feet_from_double(i * 100, 0)
                                                        Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                                        BTrecord.AppendEntity(Mtext1)
                                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                                    End If
                                                End If

                                            End If

                                        Next




                                    End If

                                    If Sta1 > Sta2 Then
                                        Dim number_of_labels As Integer = Ceiling((Sta1 - Sta2) / 100)
                                        For i = 0 To number_of_labels
                                            Dim Pt_on_Poly As Point3d
                                            If Sta1 - i * 100 >= 0 Then

                                                Pt_on_Poly = Poly2D.GetPointAtDist(Sta1 - i * 100)
                                                Dim Param_pt As Double = Poly2D.GetParameterAtPoint(Pt_on_Poly)

                                                Dim Pt_before As Point3d
                                                Dim Pt_after As Point3d
                                                If Param_pt = Round(Param_pt, 0) And Param_pt >= 1 And Param_pt <= Poly2D.NumberOfVertices - 2 Then
                                                    Pt_before = Poly2D.GetPointAtParameter(Param_pt + 1)
                                                    Pt_after = Poly2D.GetPointAtParameter(Param_pt)
                                                ElseIf Param_pt = 0 Then
                                                    Pt_before = Poly2D.GetPointAtParameter(1)
                                                    Pt_after = Poly2D.GetPointAtParameter(0)
                                                ElseIf Param_pt = Poly2D.NumberOfVertices - 1 Then
                                                    Pt_before = Poly2D.GetPointAtParameter(Param_pt)
                                                    Pt_after = Poly2D.GetPointAtParameter(Param_pt - 1)
                                                Else
                                                    Pt_before = Poly2D.GetPointAtParameter(Ceiling(Param_pt))
                                                    Pt_after = Poly2D.GetPointAtParameter(Floor(Param_pt))
                                                End If

                                                Dim Mtext_rotation As Double
                                                Mtext_rotation = GET_Bearing_rad(Pt_before.X, Pt_before.Y, Pt_after.X, Pt_after.Y)

                                                Dim Line1 As New Line(Pt_on_Poly, Pt_before)
                                                Dim Circle1 As New Circle(Pt_on_Poly, Vector3d.ZAxis, Line_length)
                                                Dim Col_int1 As New Point3dCollection
                                                Line1.IntersectWith(Circle1, Intersect.OnBothOperands, Col_int1, IntPtr.Zero, IntPtr.Zero)
                                                Dim Line2 As New Line
                                                If Col_int1.Count > 0 Then
                                                    Line2 = New Line(Pt_on_Poly, Col_int1(0))
                                                    Line2.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt_on_Poly))
                                                    Line2.TransformBy(Matrix3d.Displacement(Line2.GetPointAtDist(Line2.Length / 2).GetVectorTo(Pt_on_Poly)))
                                                    BTrecord.AppendEntity(Line2)
                                                    Trans1.AddNewlyCreatedDBObject(Line2, True)
                                                Else
                                                    Line1.IntersectWith(Circle1, Intersect.ExtendThis, Col_int1, IntPtr.Zero, IntPtr.Zero)
                                                    If Col_int1.Count > 0 Then
                                                        Line2 = New Line(Pt_on_Poly, Col_int1(0))
                                                        Line2.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Pt_on_Poly))
                                                        Line2.TransformBy(Matrix3d.Displacement(Line2.GetPointAtDist(Line2.Length / 2).GetVectorTo(Pt_on_Poly)))
                                                        BTrecord.AppendEntity(Line2)
                                                        Trans1.AddNewlyCreatedDBObject(Line2, True)
                                                    End If
                                                End If
                                                If IsNothing(Line2) = False Then
                                                    Dim Circle2 As New Circle(Pt_on_Poly, Vector3d.ZAxis, Dist_CL_to_text)
                                                    Dim Col_int2 As New Point3dCollection
                                                    Line2.IntersectWith(Circle2, Intersect.ExtendThis, Col_int2, IntPtr.Zero, IntPtr.Zero)
                                                    If Col_int2.Count > 0 Then
                                                        Dim Pt_ins As New Point3d
                                                        Pt_ins = Col_int2(0)
                                                        If Col_int2.Count = 2 Then
                                                            Pt_ins = Col_int2(1)
                                                        End If
                                                        Dim Mtext1 As New MText
                                                        Mtext1.Location = Pt_ins
                                                        Mtext1.Rotation = Mtext_rotation
                                                        Mtext1.TextHeight = Mtext_height
                                                        Mtext1.Contents = Get_chainage_feet_from_double(i * 100, 0)
                                                        Mtext1.Attachment = AttachmentPoint.MiddleCenter
                                                        BTrecord.AppendEntity(Mtext1)
                                                        Trans1.AddNewlyCreatedDBObject(Mtext1, True)
                                                    End If
                                                End If

                                            End If

                                        Next




                                    End If


                                End If



                            End If







                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If


            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub




    <CommandMethod("rot_180", CommandFlags.UsePickSet)>
    Public Sub rotate_Mtext_text_180()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Editor1.SetImpliedSelection(Empty_array)


        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select text and mtext objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If



            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        For J = 0 To Rezultat1.Value.Count - 1
                            Dim Ent1 As Entity
                            Ent1 = Trans1.GetObject(Rezultat1.Value.Item(J).ObjectId, OpenMode.ForRead)





                            If TypeOf Ent1 Is DBText Then
                                Dim Text1 As DBText = Ent1
                                Text1.UpgradeOpen()
                                Text1.Rotation = Text1.Rotation + PI
                            End If

                            If TypeOf Ent1 Is MText Then
                                Dim mText1 As MText = Ent1
                                mText1.UpgradeOpen()
                                mText1.Rotation = mText1.Rotation + PI
                            End If

                        Next




                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If


            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    <CommandMethod("PPL_DEFL_FOR_ASBUILT")>
    Public Sub DEFLECTION_FROM_3D_all_points_2_POLY()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Editor1.SetImpliedSelection(Empty_array)

        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select the stationing 3D polyline:"

            Object_Prompt.SingleOnly = True
            Rezultat1 = Editor1.GetSelection(Object_Prompt)

            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If


            Dim Rezultat2 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt2.MessageForAdding = vbLf & "Select the simplified polyline:"

            Object_Prompt2.SingleOnly = True
            Rezultat2 = Editor1.GetSelection(Object_Prompt2)


            If Rezultat2.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If


            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK And Rezultat2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False And IsNothing(Rezultat2) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2


                        Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

                        Dim Ent1 As Entity
                        Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForRead)
                        Dim Ent2 As Entity
                        Ent2 = Trans1.GetObject(Rezultat2.Value.Item(0).ObjectId, OpenMode.ForWrite)


                        If TypeOf Ent1 Is Polyline3d And (TypeOf Ent2 Is Polyline3d Or TypeOf Ent2 Is Polyline) Then
                            Dim Poly2D1 As New Polyline
                            Dim Poly3D1 As Polyline3d = Ent1
                            Dim Poly2D2 As Polyline
                            Dim Poly3D2 As Polyline3d

                            If TypeOf Ent2 Is Polyline3d Then
                                Poly3D2 = Ent2
                                Poly2D2 = New Polyline
                            End If

                            If TypeOf Ent2 Is Polyline Then
                                Poly2D2 = Ent2
                                Poly2D2.Elevation = 0
                            End If


                            Dim Index2d1 As Double = 0
                            For Each ObjId As ObjectId In Poly3D1
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D1.AddVertexAt(Index2d1, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d1 = Index2d1 + 1
                            Next
                            Poly2D1.Elevation = 0


                            If TypeOf Ent2 Is Polyline3d Then
                                Dim Index2d2 As Double = 0
                                For Each ObjId As ObjectId In Poly3D2
                                    Dim vertex2 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                    Poly2D2.AddVertexAt(Index2d2, New Point2d(vertex2.Position.X, vertex2.Position.Y), 0, 0, 0)
                                    Index2d2 = Index2d2 + 1
                                Next
                                Poly2D2.Elevation = 0
                            End If


                            W1 = get_new_worksheet_from_Excel()
                            W1.Range("A" & 1).Value = "STATION ON ORIGINAL"

                            W1.Range("C" & 1).Value = "DEFLECTION"
                            W1.Range("D" & 1).Value = "DEFLECTION DD"
                            W1.Range("E" & 1).Value = "X ORIGINAL"
                            W1.Range("F" & 1).Value = "Y ORIGINAL"
                            W1.Range("G" & 1).Value = "Z ORIGINAL"
                            W1.Range("H" & 1).Value = "X SIMPLIFIED"
                            W1.Range("I" & 1).Value = "Y SIMPLIFIED"
                            W1.Range("J" & 1).Value = "Z SIMPLIFIED"


                            For j = 0 To Poly2D2.NumberOfVertices - 1



                                If j > 0 And j < Poly2D2.NumberOfVertices - 1 Then

                                    Dim vector1 As Vector3d = Poly2D2.GetPoint3dAt(j - 1).GetVectorTo(Poly2D2.GetPoint3dAt(j))
                                    If vector1.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector1.Length < 0.01
                                            If j - K >= 0 Then
                                                vector1 = Poly2D2.GetPoint3dAt(j - K).GetVectorTo(Poly2D2.GetPoint3dAt(j))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim vector2 As Vector3d = Poly2D2.GetPoint3dAt(j).GetVectorTo(Poly2D2.GetPoint3dAt(j + 1))
                                    If vector2.Length < 0.01 Then
                                        Dim K As Double = 2
                                        Do While vector2.Length < 0.01
                                            If j + K <= Poly2D2.NumberOfVertices - 1 Then
                                                vector2 = Poly2D2.GetPoint3dAt(j).GetVectorTo(Poly2D2.GetPoint3dAt(j + K))
                                            Else
                                                Exit Do
                                            End If
                                            K = K + 1
                                        Loop
                                    End If

                                    Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                                    Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                                    Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                                    Dim Mleader1 As New MLeader
                                    Dim LT_RT As String = ""


                                    If Bearing1 < 180 Then
                                        If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                            LT_RT = " LT"
                                        Else
                                            LT_RT = " RT"
                                        End If
                                    Else
                                        If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                            LT_RT = " RT"
                                        Else
                                            LT_RT = " LT"
                                        End If
                                    End If

                                    Dim Degree As Integer = Floor(angle1)
                                    Dim Min1 As Integer = Floor((angle1 - Floor(angle1)) * 60)
                                    Dim Sec1 As Integer = Round(((angle1 - Degree) * 60 - Min1) * 60, 0)
                                    If Sec1 = 60 Then
                                        Sec1 = 0
                                        Min1 = Min1 + 1
                                    End If
                                    If Min1 = 60 Then
                                        Degree = Degree + 1
                                        Min1 = 0
                                    End If

                                    Dim Minute As String = Min1.ToString
                                    If Len(Minute) = 1 Then Minute = "0" & Minute
                                    Dim Second As String = Sec1.ToString
                                    If Len(Second) = 1 Then Second = "0" & Second

                                    Dim AngleDMS As String = Degree.ToString & "°" & Minute & "'" & Second & Chr(34)



                                    Dim continut_mleader As String = AngleDMS & LT_RT

                                    Dim Point_on_2d1 As New Point3d
                                    Point_on_2d1 = Poly2D1.GetClosestPointTo(Poly2D2.GetPointAtParameter(j), Vector3d.ZAxis, False)
                                    Dim Param_2d1 As Double = Poly2D1.GetParameterAtPoint(Point_on_2d1)


                                    W1.Range("A" & RowXL).Value = Poly3D1.GetDistanceAtParameter(Param_2d1)
                                    W1.Range("C" & RowXL).Value = continut_mleader
                                    W1.Range("D" & RowXL).Value = angle1
                                    W1.Range("E" & RowXL).Value = Poly3D1.GetPointAtParameter(Param_2d1).X
                                    W1.Range("F" & RowXL).Value = Poly3D1.GetPointAtParameter(Param_2d1).Y
                                    W1.Range("G" & RowXL).Value = Poly3D1.GetPointAtParameter(Param_2d1).Z
                                    If TypeOf Ent2 Is Polyline3d Then
                                        W1.Range("H" & RowXL).Value = Poly3D2.GetPointAtParameter(j).X
                                        W1.Range("I" & RowXL).Value = Poly3D2.GetPointAtParameter(j).Y
                                        W1.Range("J" & RowXL).Value = Poly3D2.GetPointAtParameter(j).Z
                                    End If
                                    If TypeOf Ent2 Is Polyline Then
                                        W1.Range("H" & RowXL).Value = Poly2D2.GetPointAtParameter(j).X
                                        W1.Range("I" & RowXL).Value = Poly2D2.GetPointAtParameter(j).Y
                                        W1.Range("J" & RowXL).Value = 0
                                    End If

                                    If TypeOf Ent2 Is Polyline3d Then
                                        Creaza_Mleader_nou_fara_UCS_transform(Poly3D2.GetPointAtParameter(j), Get_String_Rounded(Poly3D1.GetDistanceAtParameter(Param_2d1), 2) & vbCrLf & continut_mleader, 1, 1, 1, 1, 3)
                                    End If
                                    If TypeOf Ent2 Is Polyline Then
                                        Creaza_Mleader_nou_fara_UCS_transform(Poly2D2.GetPointAtParameter(j), Get_String_Rounded(Poly3D1.GetDistanceAtParameter(Param_2d1), 2) & vbCrLf & continut_mleader, 1, 1, 1, 1, 3)
                                    End If

                                    RowXL = RowXL + 1


                                Else
                                    Dim Point_on_2d1 As New Point3d
                                    Point_on_2d1 = Poly2D1.GetClosestPointTo(Poly2D2.GetPointAtParameter(j), Vector3d.ZAxis, False)
                                    Dim Param_2d1 As Double = Poly2D1.GetParameterAtPoint(Point_on_2d1)

                                    W1.Range("A" & RowXL).Value = Poly3D1.GetDistanceAtParameter(Param_2d1)
                                    W1.Range("E" & RowXL).Value = Poly3D1.GetPointAtParameter(Param_2d1).X
                                    W1.Range("F" & RowXL).Value = Poly3D1.GetPointAtParameter(Param_2d1).Y
                                    W1.Range("G" & RowXL).Value = Poly3D1.GetPointAtParameter(Param_2d1).Z
                                    If TypeOf Ent2 Is Polyline3d Then
                                        W1.Range("H" & RowXL).Value = Poly3D2.GetPointAtParameter(j).X
                                        W1.Range("I" & RowXL).Value = Poly3D2.GetPointAtParameter(j).Y
                                        W1.Range("J" & RowXL).Value = Poly3D2.GetPointAtParameter(j).Z
                                    End If
                                    If TypeOf Ent2 Is Polyline Then
                                        W1.Range("H" & RowXL).Value = Poly2D2.GetPointAtParameter(j).X
                                        W1.Range("I" & RowXL).Value = Poly2D2.GetPointAtParameter(j).Y
                                        W1.Range("J" & RowXL).Value = 0
                                    End If

                                    RowXL = RowXL + 1

                                End If
                            Next

                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If


            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("st_at_pt")>
    Public Sub Creaza_CHAINAGE_LABEL_ON_THE_2DPOLYLINE_1()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly1 As Polyline
            Dim Poly3D As Polyline3d

            Dim Point_on_poly As New Point3d

            Dim Dist_from_start_to_selected As Double
            Dim Station_selected As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)



                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly1 = Ent1

                            Dim Point_SELECTED As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_sel As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select the reference point:")
                            PP0.AllowNone = True
                            Point_SELECTED = Editor1.GetPoint(PP0)
                            If Not Point_SELECTED.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                                Exit Sub
                            Else
                                'aici am tratat ucs-ul 
                                Point_sel = Poly1.GetClosestPointTo(Point_SELECTED.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            End If

                            Dist_from_start_to_selected = Poly1.GetDistAtPoint(Point_sel)

                            Dim Rezultat2 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the reference station:")
                            Rezultat2.AllowNone = False
                            Dim Rezultat22 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat2)
                            If Rezultat22.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Station_selected = Rezultat22.Value
                            Else
                                Exit Sub

                            End If


                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Dim Point0 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim Point_sel As New Point3d

                            Dim PP0 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select the reference point:")
                            PP0.AllowNone = True
                            Point0 = Editor1.GetPoint(PP0)
                            If Not Point0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Exit Sub
                            Else
                                'aici am tratat ucs-ul 
                                Point_sel = Poly3D.GetClosestPointTo(Point0.Value.TransformBy(Editor1.CurrentUserCoordinateSystem), Vector3d.ZAxis, False)
                            End If

                            Dist_from_start_to_selected = Poly3D.GetDistAtPoint(Point_sel)
                            Dim Rezultat2 As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify the reference station:")
                            Rezultat2.AllowNone = False
                            Dim Rezultat22 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Rezultat2)
                            If Rezultat22.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Station_selected = Rezultat22.Value
                            Else
                                Exit Sub

                            End If






                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If
            Creaza_layer("NO PLOT", 40, "NO PLOT", False)




1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please Pick a point on the same polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim Distanta_pana_la_xing As Double
                If IsNothing(Poly1) = False Then
                    Point_on_poly = Poly1.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly1.GetDistAtPoint(Point_on_poly)
                End If

                If IsNothing(Poly3D) = False Then
                    Point_on_poly = Poly3D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)
                    Distanta_pana_la_xing = Poly3D.GetDistAtPoint(Point_on_poly)
                End If


                Dim Chainage As Double = Station_selected + Distanta_pana_la_xing - Dist_from_start_to_selected




                If Dist_from_start_to_selected + Chainage < 0 Then
                    MsgBox("The 0+000 position and your desired station point is not matching.")
                    Exit Sub
                End If




                Dim Chainage_string As String = Get_chainage_feet_from_double(Chainage, 0)
                If Chainage_string = "-0+000.0" Then Chainage_string = "0+000.0"

                Dim Mleader1 As New MLeader

                If IsNothing(Point_on_poly) = False Then
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(Point_on_poly, Chainage_string, 10, 0.1, 5, 20, 50)
                    Mleader1.Layer = "NO PLOT"
                End If


                Trans1.Commit()

                GoTo 1234

            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("xref_edit")>
    Public Sub Show_GLOBAL_CHANGE_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is Multiple_drawings_change_form Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New Multiple_drawings_change_form
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    <CommandMethod("PPL_blocks_on_prof")>
    Public Sub align_blocks_on_graph()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
        Editor1 = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Dim Curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem

        Try




            Dim RezultatCL As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_PromptCL As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the graph polyline:")

            Object_PromptCL.SetRejectMessage(vbLf & "Please select a polyline")
            Object_PromptCL.AddAllowedClass(GetType(Polyline), True)

            RezultatCL = Editor1.GetEntity(Object_PromptCL)


            If Not RezultatCL.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                MsgBox("NO polyline")
                Editor1.WriteMessage(vbLf & "Command:")

                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If







            Dim RezultatBlocks As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt_BLOCKS As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt_BLOCKS.MessageForAdding = vbLf & "Select the blocks:"

            Object_Prompt_BLOCKS.SingleOnly = False

            RezultatBlocks = Editor1.GetSelection(Object_Prompt_BLOCKS)


            If RezultatBlocks.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If



            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim Poly1 As Polyline = TryCast(Trans1.GetObject(RezultatCL.ObjectId, OpenMode.ForRead), Polyline)

                If IsNothing(Poly1) = False Then


                    For i = 0 To RezultatBlocks.Value.Count - 1
                        Dim Block1 As BlockReference = TryCast(Trans1.GetObject(RezultatBlocks.Value(i).ObjectId, OpenMode.ForRead), BlockReference)
                        If IsNothing(Block1) = False Then
                            Block1.UpgradeOpen()

                            Dim x1 As Double = Block1.Position.X
                            Dim y1 As Double = Block1.Position.Y

                            Dim Line1 As New Line(New Point3d(x1, y1 - 10000, Poly1.Elevation), New Point3d(x1, y1 + 10000, Poly1.Elevation))
                            Dim Col_int As New Point3dCollection
                            Col_int = Intersect_on_both_operands(Line1, Poly1)

                            If Col_int.Count > 0 Then
                                Block1.Position = Col_int(0)


                            End If


                        End If


                    Next
                End If
                Trans1.Commit()
            End Using




            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("adjust_elev")>
    Public Sub modify_z_of_a_popyline_vertex()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Editor1.SetImpliedSelection(Empty_array)

        Try
            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select 3D polyline:"

            Object_Prompt.SingleOnly = True
            Rezultat1 = Editor1.GetSelection(Object_Prompt)

            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If





            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
                        Dim RowXL As Integer = 2



                        Dim Ent1 As Entity
                        Ent1 = Trans1.GetObject(Rezultat1.Value.Item(0).ObjectId, OpenMode.ForWrite)



                        If TypeOf Ent1 Is Polyline3d Then

                            Dim Poly3D1 As Polyline3d = Ent1



                            Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick the node for adjust:")

                            PP1.AllowNone = False
                            Result_point1 = Editor1.GetPoint(PP1)
                            If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then

                                Exit Sub
                            End If

                            Dim Result_point2 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP2 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Please pick the point1:")

                            PP2.AllowNone = False
                            Result_point2 = Editor1.GetPoint(PP2)
                            If Result_point2.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then

                                Exit Sub
                            End If

                            Dim Result_point3 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                            Dim PP3 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick the point2:")

                            PP3.AllowNone = False
                            Result_point3 = Editor1.GetPoint(PP3)
                            If Result_point3.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then

                                Exit Sub
                            End If

                            Dim Pt_on_poly1 As New Point3d

                            Pt_on_poly1 = Poly3D1.GetClosestPointTo(Result_point1.Value, Vector3d.ZAxis, False)

                            Dim new_z As Double = 0
                            Dim l1 As Double = ((Result_point2.Value.X - Result_point3.Value.X) ^ 2 + (Result_point2.Value.Y - Result_point3.Value.Y) ^ 2) ^ 0.5
                            Dim d1 As Double = ((Result_point2.Value.X - Pt_on_poly1.X) ^ 2 + (Result_point2.Value.Y - Pt_on_poly1.Y) ^ 2) ^ 0.5
                            Dim DeltaZ As Double = Result_point3.Value.Z - Result_point2.Value.Z
                            new_z = Result_point2.Value.Z + DeltaZ * d1 / l1

                            Dim vertex_adjustm As PolylineVertex3d

                            For Each ObjId As ObjectId In Poly3D1
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                Dim Dist As Double = ((Pt_on_poly1.X - vertex1.Position.X) ^ 2 + (Pt_on_poly1.Y - vertex1.Position.Y) ^ 2) ^ 0.5
                                If Dist < 0.1 Then
                                    vertex1.Position = New Point3d(vertex1.Position.X, vertex1.Position.Y, new_z)
                                End If


                            Next





                        End If

                        Trans1.Commit()
                    End Using

                Else

                    Exit Sub
                End If
            End If


            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("FIND_DUPLICATES")>
    Public Sub fIND_duplicates()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId
        Editor1.SetImpliedSelection(Empty_array)

        Try





            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                Dim List1 As New List(Of ObjectId)

                For Each Objid As ObjectId In BTrecord
                    Dim Ent1 As Entity
                    Ent1 = Trans1.GetObject(Objid, OpenMode.ForRead)
                    List1.Add(Objid)
                    If TypeOf Ent1 Is MText Then

                        Dim Mtext1 As MText = Ent1
                        Dim Continut1 As String = Mtext1.Text
                        For Each Objid2 As ObjectId In BTrecord
                            If List1.Contains(Objid2) = False Then
                                Dim Ent2 As Entity
                                Ent2 = Trans1.GetObject(Objid2, OpenMode.ForRead)
                                If TypeOf Ent2 Is MText Then
                                    Dim Mtext2 As MText = Ent2
                                    Dim Continut2 As String = Mtext2.Text
                                    If Continut1 = Continut2 Then
                                        Mtext2.UpgradeOpen()
                                        Mtext2.ColorIndex = 1


                                    End If



                                End If
                            End If





                        Next
                    End If
                Next







                Trans1.Commit()
            End Using




            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("survey_match")>
    Public Sub Show_survey_tools_form()
        If isSECURE() = False Then Exit Sub
        For Each forma In System.Windows.Forms.Application.OpenForms
            If TypeOf forma Is SURVEY_TOOLS_FORM Then
                forma.Focus()
                forma.WindowState = Windows.Forms.FormWindowState.Normal
                Exit Sub
            End If
        Next
        Try
            Dim forma1 As New SURVEY_TOOLS_FORM
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    <CommandMethod("DEFLECTION_ALL", CommandFlags.UsePickSet)>
    Public Sub DEFLECTION_FROM_3D_all_points()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId

        Try


            Dim Rezultat_poly As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the polyline:")

            Object_Prompt1.SetRejectMessage(vbLf & "Please select a 3d polyline or a polyline")
            Object_Prompt1.AddAllowedClass(GetType(Polyline3d), True)
            Object_Prompt1.AddAllowedClass(GetType(Polyline), True)

            Rezultat_poly = Editor1.GetEntity(Object_Prompt1)


            If Not Rezultat_poly.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                MsgBox("NO centerline")
                Editor1.WriteMessage(vbLf & "Command:")

                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If

            Dim dt1 As New System.Data.DataTable
            dt1.Columns.Add("measured_distance", GetType(Double))
            dt1.Columns.Add("x", GetType(Double))
            dt1.Columns.Add("y", GetType(Double))
            dt1.Columns.Add("z", GetType(Double))
            dt1.Columns.Add("deflection", GetType(String))
            dt1.Columns.Add("deflection_dd", GetType(Double))

            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)



                Dim Planul_curent As Plane = New Plane(New Point3d(0, 0, 0), Vector3d.ZAxis)

                Dim Ent1 As Entity = Trans1.GetObject(Rezultat_poly.ObjectId, OpenMode.ForRead)



                If TypeOf Ent1 Is Polyline3d Then



                    Dim Poly2D As New Polyline
                    Dim Poly3D As Polyline3d = Ent1

                    Dim Index2d As Double = 0
                    For Each ObjId As ObjectId In Poly3D
                        Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                        Index2d = Index2d + 1
                    Next
                    Poly2D.Elevation = 0






                    For j = 0 To Poly2D.NumberOfVertices - 1



                        If j > 0 And j < Poly2D.NumberOfVertices - 1 Then

                            Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                            If vector1.Length < 0.01 Then
                                Dim K As Double = 2
                                Do While vector1.Length < 0.01
                                    If j - K >= 0 Then
                                        vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    Else
                                        Exit Do
                                    End If
                                    K = K + 1
                                Loop
                            End If

                            Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                            If vector2.Length < 0.01 Then
                                Dim K As Double = 2
                                Do While vector2.Length < 0.01
                                    If j + K <= Poly2D.NumberOfVertices - 1 Then
                                        vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                    Else
                                        Exit Do
                                    End If
                                    K = K + 1
                                Loop
                            End If

                            Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                            Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                            Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                            Dim Mleader1 As New MLeader
                            Dim LT_RT As String = ""


                            If Bearing1 < 180 Then
                                If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                    LT_RT = " LT"
                                Else
                                    LT_RT = " RT"
                                End If
                            Else
                                If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                    LT_RT = " RT"
                                Else
                                    LT_RT = " LT"
                                End If
                            End If

                            Dim Degree As Integer = Floor(angle1)
                            Dim Min1 As Integer = Floor((angle1 - Floor(angle1)) * 60)
                            Dim Sec1 As Integer = Round(((angle1 - Degree) * 60 - Min1) * 60, 0)
                            If Sec1 = 60 Then
                                Sec1 = 0
                                Min1 = Min1 + 1
                            End If
                            If Min1 = 60 Then
                                Degree = Degree + 1
                                Min1 = 0
                            End If

                            Dim Minute As String = Min1.ToString
                            If Len(Minute) = 1 Then Minute = "0" & Minute
                            Dim Second As String = Sec1.ToString
                            If Len(Second) = 1 Then Second = "0" & Second

                            Dim AngleDMS As String = Degree.ToString & "°" & Minute & "'" & Second & Chr(34)



                            Dim deflection_dms As String = AngleDMS & LT_RT

                            dt1.Rows.Add()
                            dt1.Rows(dt1.Rows.Count - 1).Item("measured_distance") = Poly3D.GetDistanceAtParameter(j)
                            dt1.Rows(dt1.Rows.Count - 1).Item("x") = Poly3D.GetPointAtParameter(j).X
                            dt1.Rows(dt1.Rows.Count - 1).Item("y") = Poly3D.GetPointAtParameter(j).Y
                            dt1.Rows(dt1.Rows.Count - 1).Item("z") = Poly3D.GetPointAtParameter(j).Z
                            dt1.Rows(dt1.Rows.Count - 1).Item("deflection") = deflection_dms
                            dt1.Rows(dt1.Rows.Count - 1).Item("deflection_dd") = angle1


                        Else
                            dt1.Rows.Add()
                            dt1.Rows(dt1.Rows.Count - 1).Item("measured_distance") = Poly3D.GetDistanceAtParameter(j)
                            dt1.Rows(dt1.Rows.Count - 1).Item("x") = Poly3D.GetPointAtParameter(j).X
                            dt1.Rows(dt1.Rows.Count - 1).Item("y") = Poly3D.GetPointAtParameter(j).Y
                            dt1.Rows(dt1.Rows.Count - 1).Item("z") = Poly3D.GetPointAtParameter(j).Z

                        End If
                    Next

                End If

                If TypeOf Ent1 Is Polyline Then



                    Dim Poly2D As Polyline = Ent1

                    For j = 0 To Poly2D.NumberOfVertices - 1



                        If j > 0 And j < Poly2D.NumberOfVertices - 1 Then

                            Dim vector1 As Vector3d = Poly2D.GetPoint3dAt(j - 1).GetVectorTo(Poly2D.GetPoint3dAt(j))
                            If vector1.Length < 0.01 Then
                                Dim K As Double = 2
                                Do While vector1.Length < 0.01
                                    If j - K >= 0 Then
                                        vector1 = Poly2D.GetPoint3dAt(j - K).GetVectorTo(Poly2D.GetPoint3dAt(j))
                                    Else
                                        Exit Do
                                    End If
                                    K = K + 1
                                Loop
                            End If

                            Dim vector2 As Vector3d = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + 1))
                            If vector2.Length < 0.01 Then
                                Dim K As Double = 2
                                Do While vector2.Length < 0.01
                                    If j + K <= Poly2D.NumberOfVertices - 1 Then
                                        vector2 = Poly2D.GetPoint3dAt(j).GetVectorTo(Poly2D.GetPoint3dAt(j + K))
                                    Else
                                        Exit Do
                                    End If
                                    K = K + 1
                                Loop
                            End If

                            Dim Bearing1 As Double = (vector1.AngleOnPlane(Planul_curent)) * 180 / PI
                            Dim Bearing2 As Double = (vector2.AngleOnPlane(Planul_curent)) * 180 / PI

                            Dim angle1 As Double = (vector2.GetAngleTo(vector1)) * 180 / PI
                            Dim Mleader1 As New MLeader
                            Dim LT_RT As String = ""


                            If Bearing1 < 180 Then
                                If Bearing2 < Bearing1 + 180 And Bearing2 > Bearing1 Then
                                    LT_RT = " LT"
                                Else
                                    LT_RT = " RT"
                                End If
                            Else
                                If Bearing2 < Bearing1 And Bearing2 > Bearing1 - 180 Then
                                    LT_RT = " RT"
                                Else
                                    LT_RT = " LT"
                                End If
                            End If

                            Dim Degree As Integer = Floor(angle1)
                            Dim Min1 As Integer = Floor((angle1 - Floor(angle1)) * 60)
                            Dim Sec1 As Integer = Round(((angle1 - Degree) * 60 - Min1) * 60, 0)
                            If Sec1 = 60 Then
                                Sec1 = 0
                                Min1 = Min1 + 1
                            End If
                            If Min1 = 60 Then
                                Degree = Degree + 1
                                Min1 = 0
                            End If

                            Dim Minute As String = Min1.ToString
                            If Len(Minute) = 1 Then Minute = "0" & Minute
                            Dim Second As String = Sec1.ToString
                            If Len(Second) = 1 Then Second = "0" & Second

                            Dim AngleDMS As String = Degree.ToString & "°" & Minute & "'" & Second & Chr(34)



                            Dim deflection_dms As String = AngleDMS & LT_RT

                            dt1.Rows.Add()
                            dt1.Rows(dt1.Rows.Count - 1).Item("measured_distance") = Poly2D.GetDistanceAtParameter(j)
                            dt1.Rows(dt1.Rows.Count - 1).Item("x") = Poly2D.GetPointAtParameter(j).X
                            dt1.Rows(dt1.Rows.Count - 1).Item("y") = Poly2D.GetPointAtParameter(j).Y
                            dt1.Rows(dt1.Rows.Count - 1).Item("z") = Poly2D.GetPointAtParameter(j).Z
                            dt1.Rows(dt1.Rows.Count - 1).Item("deflection") = deflection_dms
                            dt1.Rows(dt1.Rows.Count - 1).Item("deflection_dd") = angle1



                        Else
                            dt1.Rows.Add()
                            dt1.Rows(dt1.Rows.Count - 1).Item("measured_distance") = Poly2D.GetDistanceAtParameter(j)
                            dt1.Rows(dt1.Rows.Count - 1).Item("x") = Poly2D.GetPointAtParameter(j).X
                            dt1.Rows(dt1.Rows.Count - 1).Item("y") = Poly2D.GetPointAtParameter(j).Y
                            dt1.Rows(dt1.Rows.Count - 1).Item("z") = Poly2D.GetPointAtParameter(j).Z

                        End If
                    Next

                End If

                Trans1.Commit()
            End Using

            Transfer_datatable_to_new_excel_spreadsheet(dt1)

            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("OffUSER", CommandFlags.UsePickSet)>
    Public Sub OFFSET_WITH_USER_INPUT()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim Empty_array() As ObjectId

        Try


            Dim Rezultat_poly As Autodesk.AutoCAD.EditorInput.PromptEntityResult

            Dim Object_Prompt1 As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select the polyline:")

            Object_Prompt1.SetRejectMessage(vbLf & "Please select a 3d polyline or a polyline")
            Object_Prompt1.AddAllowedClass(GetType(Polyline3d), True)


            Rezultat_poly = Editor1.GetEntity(Object_Prompt1)


            If Not Rezultat_poly.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                MsgBox("NO centerline")
                Editor1.WriteMessage(vbLf & "Command:")

                Editor1.SetImpliedSelection(Empty_array)
                Exit Sub
            End If








1234:



            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)



                Dim Result_point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please pick position:")

                PP1.AllowNone = False
                Result_point1 = Editor1.GetPoint(PP1)
                If Result_point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel Then
                    Trans1.Commit()
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If



                Dim String_op As Autodesk.AutoCAD.EditorInput.PromptStringOptions
                String_op = New Autodesk.AutoCAD.EditorInput.PromptStringOptions(vbLf & "Specify offset and direction")
                String_op.AllowSpaces = False
                String_op.UseDefaultValue = True
                String_op.DefaultValue = "0L"

                Dim SSS As Autodesk.AutoCAD.EditorInput.PromptResult = Editor1.GetString(String_op)

                If SSS.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Trans1.Commit()
                    Editor1.WriteMessage(vbLf & "Command:")
                    Exit Sub
                End If


                Dim Str_LR As String = SSS.StringResult

                If Str_LR.Length > 1 Then
                    Dim LR As String = Right(Str_LR, 1).ToUpper

                    Dim Off1 As String = Left(Str_LR, Len(Str_LR) - 1)

                    Dim Offset_no As Double = 0

                    If IsNumeric(Off1) = True Then
                        Offset_no = CDbl(Off1)
                    End If

                    If Offset_no > 0 And (LR.ToUpper = "R" Or LR.ToUpper = "L") Then
                        Dim Ent1 As Entity = Trans1.GetObject(Rezultat_poly.ObjectId, OpenMode.ForRead)
                        Dim Poly2D As New Polyline
                        Dim Poly3D As Polyline3d = Ent1

                        Dim Index2d As Double = 0
                        For Each ObjId As ObjectId In Poly3D
                            Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                            Index2d = Index2d + 1
                        Next
                        Poly2D.Elevation = 0

                        Dim Point_on_2d As New Point3d
                        Point_on_2d = Poly2D.GetClosestPointTo(Result_point1.Value, Vector3d.ZAxis, False)

                        Dim Param0 As Double = Poly2D.GetParameterAtPoint(Point_on_2d)
                        Dim Param1 As Double = Floor(Param0)
                        Dim Param2 As Double = Ceiling(Param0)
                        If Param1 = Param2 Then
                            Param2 = Param1 + 1
                        End If


                        Dim Line1 As New Line(Poly2D.GetPointAtParameter(Floor(Param1)), Poly2D.GetPointAtParameter(Ceiling(Param2)))
                        Line1.TransformBy(Matrix3d.Displacement(Line1.StartPoint.GetVectorTo(Point_on_2d)))
                        If LR.ToUpper = "L" Then
                            Line1.TransformBy(Matrix3d.Rotation(PI / 2, Vector3d.ZAxis, Point_on_2d))
                        End If

                        If LR.ToUpper = "R" Then
                            Line1.TransformBy(Matrix3d.Rotation(-PI / 2, Vector3d.ZAxis, Point_on_2d))
                        End If

                        Line1.TransformBy(Matrix3d.Scaling(Offset_no / Line1.Length, Point_on_2d))

                        BTrecord.AppendEntity(Line1)
                        Trans1.AddNewlyCreatedDBObject(Line1, True)
                        Dim Elev_op As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify Depth:")
                        Elev_op.AllowNone = False
                        Elev_op.UseDefaultValue = True
                        Elev_op.DefaultValue = 0
                        Dim Elev1 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(Elev_op)
                        Dim Z As Double = 0

                        If Elev1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            Z = Elev1.Value
                        End If

                        Dim Pt3d As New DBPoint(New Point3d(Line1.EndPoint.X, Line1.EndPoint.Y, -Z))
                        BTrecord.AppendEntity(Pt3d)
                        Trans1.AddNewlyCreatedDBObject(Pt3d, True)
                        Trans1.Commit()
                    End If
                End If










            End Using
            GoTo 1234


            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception

            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("chatpt")>
    Public Sub Creaza_CHAINAGE_identifica_pozitia_pe_polyline()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try





            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly2D As Polyline
            Dim Poly3D As Polyline3d



            Dim Dist_ref As Double

            Dim Station_ref As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Dim Point_ref As New Point3d

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1





                            Dim R_station As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify a known station value:")
                            R_station.AllowNone = False
                            R_station.UseDefaultValue = True
                            R_station.DefaultValue = 0
                            Dim R_station1 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(R_station)
                            Station_ref = R_station1.Value

                            If R_station1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Trans1.Commit()
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If


                            Dim point_pr0 As Autodesk.AutoCAD.EditorInput.PromptPointResult


                            Dim PP_ref As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select the known point")
                            PP_ref.AllowNone = True
                            point_pr0 = Editor1.GetPoint(PP_ref)
                            If Not point_pr0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If


                            Point_ref = Poly2D.GetClosestPointTo(point_pr0.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)

                            Dist_ref = Poly2D.GetDistAtPoint(Point_ref)




                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Poly2D = New Polyline


                            Dim Index2d As Double = 0
                            For Each ObjId As ObjectId In Poly3D
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d = Index2d + 1
                            Next
                            Poly2D.Elevation = 0


                            Dim R_station As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify a known station value:")
                            R_station.AllowNone = False
                            R_station.UseDefaultValue = True
                            R_station.DefaultValue = 0
                            Dim R_station1 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(R_station)
                            Station_ref = R_station1.Value

                            Dim point_pr0 As Autodesk.AutoCAD.EditorInput.PromptPointResult


                            Dim PP_ref As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select the known point")
                            PP_ref.AllowNone = True
                            point_pr0 = Editor1.GetPoint(PP_ref)
                            If Not point_pr0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If


                            Dim Pt_2d As New Point3d

                            Pt_2d = Poly2D.GetClosestPointTo(point_pr0.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)

                            Point_ref = Poly3D.GetPointAtParameter(Poly2D.GetParameterAtPoint(Pt_2d))


                            Dist_ref = Poly3D.GetDistAtPoint(Point_ref)


                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If

1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                Dim PP1 As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbCrLf & "Please Pick a point on the same polyline:")
                Dim Point1 As Autodesk.AutoCAD.EditorInput.PromptPointResult
                PP1.AllowNone = False
                Point1 = Editor1.GetPoint(PP1)
                If Not Point1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    Editor1.SetImpliedSelection(Empty_array)
                    Trans1.Commit()
                    Exit Sub
                End If

                Dim New_pt As New Point3d



                If IsNothing(Poly3D) = False Then
                    Dim Pt2d As New Point3d
                    Pt2d = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)

                    New_pt = Poly3D.GetPointAtParameter(Poly2D.GetParameterAtPoint(Pt2d))

                    Dim Dist1 As Double = Poly3D.GetDistAtPoint(New_pt)
                    Dim new_sta As Double = Station_ref + (Dist1 - Dist_ref)
                    Dim Mleader1 As New MLeader
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(New_pt, Get_chainage_from_double(new_sta, 2), 5, 2.5, 5, 10, 10)

                    Trans1.Commit()

                    GoTo 1234
                End If

                If IsNothing(Poly2D) = False Then

                    New_pt = Poly2D.GetClosestPointTo(Point1.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)

                    Dim Dist1 As Double = Poly2D.GetDistAtPoint(New_pt)
                    Dim new_sta As Double = Station_ref + (Dist1 - Dist_ref)
                    Dim Mleader1 As New MLeader
                    Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(New_pt, Get_chainage_from_double(new_sta, 2), 5, 2.5, 5, 10, 10)

                    Trans1.Commit()

                    GoTo 1234

                End If





            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("ptatch")>
    Public Sub Creaza_point_from_chainage()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try





            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

            Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
            Object_Prompt.MessageForAdding = vbLf & "Select polyline:"

            Object_Prompt.SingleOnly = True

            Rezultat1 = Editor1.GetSelection(Object_Prompt)


            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Editor1.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If
            Dim Poly2D As Polyline
            Dim Poly3D As Polyline3d




            Dim Dist_ref As Double

            Dim Station_ref As Double

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then

                    Dim Point_ref As New Point3d

                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                        Obj1 = Rezultat1.Value.Item(0)
                        Dim Ent1 As Entity
                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then


                            Poly2D = Ent1





                            Dim R_station As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify a known station value:")
                            R_station.AllowNone = False
                            R_station.UseDefaultValue = True
                            R_station.DefaultValue = 0
                            Dim R_station1 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(R_station)
                            Station_ref = R_station1.Value

                            Dim point_pr0 As Autodesk.AutoCAD.EditorInput.PromptPointResult


                            Dim PP_ref As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select the known point")
                            PP_ref.AllowNone = True
                            point_pr0 = Editor1.GetPoint(PP_ref)
                            If Not point_pr0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If


                            Point_ref = Poly2D.GetClosestPointTo(point_pr0.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)

                            Dist_ref = Poly2D.GetDistAtPoint(Point_ref)




                            Trans1.Commit()

                        ElseIf TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then


                            Poly3D = Ent1

                            Poly2D = New Polyline


                            Dim Index2d As Double = 0
                            For Each ObjId As ObjectId In Poly3D
                                Dim vertex1 As PolylineVertex3d = Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                Poly2D.AddVertexAt(Index2d, New Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0)
                                Index2d = Index2d + 1
                            Next
                            Poly2D.Elevation = 0


                            Dim R_station As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify a known station value:")
                            R_station.AllowNone = False
                            R_station.UseDefaultValue = True
                            R_station.DefaultValue = 0
                            Dim R_station1 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(R_station)
                            Station_ref = R_station1.Value

                            Dim point_pr0 As Autodesk.AutoCAD.EditorInput.PromptPointResult


                            Dim PP_ref As New Autodesk.AutoCAD.EditorInput.PromptPointOptions(vbLf & "Select the known point")
                            PP_ref.AllowNone = True
                            point_pr0 = Editor1.GetPoint(PP_ref)
                            If Not point_pr0.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                Editor1.WriteMessage(vbLf & "Command:")
                                Exit Sub
                            End If


                            Dim Pt_2d As New Point3d

                            Pt_2d = Poly2D.GetClosestPointTo(point_pr0.Value.TransformBy(curent_ucs_matrix), Vector3d.ZAxis, False)

                            Point_ref = Poly3D.GetPointAtParameter(Poly2D.GetParameterAtPoint(Pt_2d))


                            Dist_ref = Poly3D.GetDistAtPoint(Point_ref)


                            Trans1.Commit()

                        Else
                            Editor1.WriteMessage("No Polyline")
                            Editor1.WriteMessage(vbLf & "Command:")
                            Exit Sub
                        End If
                    End Using
                End If
            End If




1234:
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)



                Try
                    Dim R_station As New Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(vbLf & "Specify a new station value:")
                    R_station.AllowNone = False
                    R_station.UseDefaultValue = True
                    R_station.DefaultValue = -1
                    Dim R_station1 As Autodesk.AutoCAD.EditorInput.PromptDoubleResult = Editor1.GetDouble(R_station)

                    If R_station1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Trans1.Commit()
                        Editor1.WriteMessage(vbLf & "Command:")
                        Exit Sub
                    End If

                    Dim Station1 As Double = R_station1.Value

                    Dim New_pt As New Point3d



                    If IsNothing(Poly3D) = False Then


                        New_pt = Poly3D.GetPointAtDist(Dist_ref + (Station1 - Station_ref))


                        Dim Mleader1 As New MLeader
                        Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(New_pt, Get_chainage_from_double(Station1, 2), 5, 2.5, 5, 10, 10)

                        Trans1.Commit()

                        GoTo 1234
                    End If

                    If IsNothing(Poly2D) = False Then


                        New_pt = Poly2D.GetPointAtDist(Dist_ref + (Station1 - Station_ref))


                        Dim Mleader1 As New MLeader
                        Mleader1 = Creaza_Mleader_nou_fara_UCS_transform(New_pt, Get_chainage_from_double(Station1, 2), 5, 2.5, 5, 10, 10)

                        Trans1.Commit()

                        GoTo 1234

                    End If
                Catch ex As System.Exception
                    Trans1.Commit()
                    Editor1.SetImpliedSelection(Empty_array)
                    Editor1.WriteMessage(vbLf & "Command:")
                    MsgBox(ex.Message)
                End Try







            End Using


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub

    <CommandMethod("ATTSYNC_GLOBAL")>
    Public Sub ATT_SYNC_GLOBAL()
        If isSECURE() = False Then Exit Sub
        Dim Empty_array() As ObjectId
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor
        Dim curent_ucs_matrix As Matrix3d = Editor1.CurrentUserCoordinateSystem
        Try

            Dim lista1 As New List(Of String)
            Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                Dim BlockTable1 As BlockTable = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                For Each id1 As ObjectId In BlockTable1
                    Dim btr As BlockTableRecord = Trans1.GetObject(id1, OpenMode.ForRead)
                    If btr.IsDynamicBlock = True Or btr.HasAttributeDefinitions = True Then
                        lista1.Add(btr.Name)
                    End If
                Next
            End Using

            If lista1.Count > 0 Then
                For i = 0 To lista1.Count - 1
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("ATTSYNC" & vbCr & "N" & vbCr & lista1.Item(i) & vbCr, False, False, True)
                Next
            End If


            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
        Catch ex As Exception
            Editor1.SetImpliedSelection(Empty_array)
            Editor1.WriteMessage(vbLf & "Command:")
            MsgBox(ex.Message)
        End Try



    End Sub






    <CommandMethod("asif1", CommandFlags.UsePickSet)>
    Public Sub write_special_to_excel()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet


                        W1 = Get_active_worksheet_from_Excel()

                        Dim is_poly2d As Boolean = False
                        Dim is_block As Boolean = False

                        Dim label1 As String = ""
                        Dim layer1 As String = ""
                        Dim pt1 As New Point3d()
                        If Rezultat1.Value.Count > 0 Then
                            If RowXL = 1 Then

                                W1.Range("A" & RowXL).Value = "Label"
                                W1.Range("B" & RowXL).Value = "X"
                                W1.Range("C" & RowXL).Value = "Y"
                                W1.Range("D" & RowXL).Value = "Layer"
                                RowXL = RowXL + 1
                            End If

                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForWrite)
                                Creaza_layer(Ent1.Layer & "_solved", 7, "", False)


                                If TypeOf Ent1 Is Polyline Then
                                    Dim poly1 As Polyline = Ent1
                                    If poly1.NumberOfVertices = 2 Then
                                        pt1 = poly1.GetPointAtDist(poly1.Length / 2)
                                    End If
                                    If poly1.NumberOfVertices = 11 Then
                                        Dim point1 As Point2d = poly1.GetPoint2dAt(1)
                                        Dim point6 As Point2d = poly1.GetPoint2dAt(6)
                                        Dim point3 As Point2d = poly1.GetPoint2dAt(3)
                                        Dim point8 As Point2d = poly1.GetPoint2dAt(8)

                                        Dim polyt1 = New Polyline
                                        polyt1.AddVertexAt(0, point1, 0, 0, 0)
                                        polyt1.AddVertexAt(1, point6, 0, 0, 0)

                                        Dim polyt2 = New Polyline
                                        polyt2.AddVertexAt(0, point3, 0, 0, 0)
                                        polyt2.AddVertexAt(1, point8, 0, 0, 0)

                                        Dim colint As New Point3dCollection
                                        polyt1.IntersectWith(polyt2, intersectType:=Intersect.OnBothOperands, colint, IntPtr.Zero, IntPtr.Zero)
                                        If colint.Count > 0 Then
                                            pt1 = colint(0)
                                        End If
                                    End If
                                    layer1 = poly1.Layer
                                    poly1.Layer = Ent1.Layer & "_solved"
                                End If



                                If TypeOf Ent1 Is BlockReference Then
                                    Dim block1 As BlockReference = Ent1
                                    pt1 = block1.Position
                                    layer1 = block1.Layer
                                    block1.Layer = Ent1.Layer & "_solved"
                                End If


                                If TypeOf Ent1 Is DBText Then
                                    Dim text1 As DBText = Ent1
                                    If label1 = "" Then
                                        label1 = text1.TextString
                                    Else
                                        label1 = label1 & vbCrLf & text1.TextString
                                    End If
                                    text1.Layer = Ent1.Layer & "_solved"
                                End If

                            Next

                            W1.Range("A" & RowXL).Value = label1
                            W1.Range("B" & RowXL).Value = pt1.X
                            W1.Range("C" & RowXL).Value = pt1.Y
                            W1.Range("D" & RowXL).Value = layer1

                            RowXL = RowXL + 1
                        End If

                        Trans1.Commit()
                    End Using
                Else
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    <CommandMethod("asif_star", CommandFlags.UsePickSet)>
    Public Sub write_star_to_excel()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet


                        W1 = Get_active_worksheet_from_Excel()

                        Dim is_poly2d As Boolean = False
                        Dim is_block As Boolean = False

                        Dim label1 As String = ""
                        Dim layer1 As String = ""
                        Dim pt1 As New Point3d()
                        If Rezultat1.Value.Count > 0 Then
                            If RowXL = 1 Then

                                W1.Range("A" & RowXL).Value = "Label"
                                W1.Range("B" & RowXL).Value = "X"
                                W1.Range("C" & RowXL).Value = "Y"
                                W1.Range("D" & RowXL).Value = "Layer"
                                RowXL = RowXL + 1
                            End If
                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)


                                If TypeOf Ent1 Is Polyline Then
                                    Dim poly1 As Polyline = Ent1
                                    If poly1.NumberOfVertices = 11 Then
                                        Dim point1 As Point2d = poly1.GetPoint2dAt(1)
                                        Dim point6 As Point2d = poly1.GetPoint2dAt(6)
                                        Dim point3 As Point2d = poly1.GetPoint2dAt(3)
                                        Dim point8 As Point2d = poly1.GetPoint2dAt(8)

                                        Dim polyt1 = New Polyline
                                        polyt1.AddVertexAt(0, point1, 0, 0, 0)
                                        polyt1.AddVertexAt(1, point6, 0, 0, 0)

                                        Dim polyt2 = New Polyline
                                        polyt2.AddVertexAt(0, point3, 0, 0, 0)
                                        polyt2.AddVertexAt(1, point8, 0, 0, 0)

                                        Dim colint As New Point3dCollection
                                        polyt1.IntersectWith(polyt2, intersectType:=Intersect.OnBothOperands, colint, IntPtr.Zero, IntPtr.Zero)
                                        If colint.Count > 0 Then
                                            pt1 = colint(0)
                                            label1 = "Star"
                                            layer1 = poly1.Layer
                                            W1.Range("A" & RowXL).Value = label1
                                            W1.Range("B" & RowXL).Value = pt1.X
                                            W1.Range("C" & RowXL).Value = pt1.Y
                                            W1.Range("D" & RowXL).Value = layer1

                                            RowXL = RowXL + 1
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End Using
                Else
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub


    <CommandMethod("asif_blocks", CommandFlags.UsePickSet)>
    Public Sub write_aSIF_BLOCKS_to_excel()
        If isSECURE() = False Then Exit Sub
        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor


        Try


            Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
            Rezultat1 = Editor1.SelectImplied

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

            Else
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)
            End If



            If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")
                Exit Sub
            End If

            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                If IsNothing(Rezultat1) = False Then


                    Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                        Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                        Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                        BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                        BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet


                        W1 = Get_active_worksheet_from_Excel()

                        Dim is_poly2d As Boolean = False
                        Dim is_block As Boolean = False

                        Dim label1 As String = ""
                        Dim layer1 As String = ""
                        Dim pt1 As New Point3d()
                        If Rezultat1.Value.Count > 0 Then
                            If RowXL = 1 Then

                                W1.Range("A" & RowXL).Value = "Label"
                                W1.Range("B" & RowXL).Value = "X"
                                W1.Range("C" & RowXL).Value = "Y"
                                W1.Range("D" & RowXL).Value = "Layer"
                                RowXL = RowXL + 1
                            End If
                            For i = 0 To Rezultat1.Value.Count - 1
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent1 Is BlockReference Then
                                    Dim block1 As BlockReference = Ent1
                                    pt1 = block1.Position
                                    layer1 = block1.Layer
                                    label1 = get_block_name(block1)

                                    W1.Range("A" & RowXL).Value = label1
                                    W1.Range("B" & RowXL).Value = pt1.X
                                    W1.Range("C" & RowXL).Value = pt1.Y
                                    W1.Range("D" & RowXL).Value = layer1

                                    RowXL = RowXL + 1
                                End If

                            Next

                        End If


                    End Using
                Else
                    Exit Sub
                End If
            End If

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

        Catch ex As Exception
            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)
            MsgBox(ex.Message)
        End Try
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    <CommandMethod("SET_XLROW", CommandFlags.UsePickSet)>
    Public Sub SET_excel_row()
        RowXL = InputBox("specify xl row")


    End Sub

End Class
