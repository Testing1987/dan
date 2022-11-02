Public Class Coordinate_systems_in_Excel


    Private Sub Button_convert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_convert.Click

        Button_convert.Visible = False
        Dim Acmap As Autodesk.Gis.Map.Platform.AcMapMap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap

        Dim Curent_system As String = Acmap.GetMapSRS()

        Dim String_UTM83_12 As String = "PROJCS[" & Chr(34) & "UTM83-12" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_UTM83_11 As String = "PROJCS[" & Chr(34) & "UTM83-11" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_CANA83_10TM115 As String = "PROJCS[" & Chr(34) & "CANA83-10TM115" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999200000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-115.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_CANA83_3TM114 As String = "PROJCS[" & Chr(34) & "CANA83-3TM114" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-114.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_CANA83_3TM111 As String = "PROJCS[" & Chr(34) & "CANA83-3TM111" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_CANA83_3TM117 As String = "PROJCS[" & Chr(34) & "CANA83-3TM117" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_CANA83_3TM120 As String = "PROJCS[" & Chr(34) & "CANA83-3TM120" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.999900000000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-120.00000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.00000000000000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_UTM27_11 As String = "PROJCS[" & Chr(34) & "UTM27-11" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-117.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_UTM27_12 As String = "PROJCS[" & Chr(34) & "UTM27-12" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-111.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_LL84 As String = "GEOGCS[" & Chr(34) & "LL84" & Chr(34) & ",DATUM[" & Chr(34) & "WGS84" & Chr(34) & ",SPHEROID[" & Chr(34) & "WGS84" & Chr(34) & ",6378137.000,298.25722293]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"
        Dim String_LL83 As String = "GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.01745329251994]]"
        Dim String_UTM83_10 As String = "PROJCS[" & Chr(34) & "UTM83-10" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-123.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_UTM27_10 As String = "PROJCS[" & Chr(34) & "UTM27-10" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL27" & Chr(34) & ",DATUM[" & Chr(34) & "NAD27" & Chr(34) & ",SPHEROID[" & Chr(34) & "CLRK66" & Chr(34) & ",6378206.400,294.97869821]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-123.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim String_UTM83_9 As String = "PROJCS[" & Chr(34) & "UTM83-9" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",500000.000],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-129.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Meter" & Chr(34) & ",1.00000000000000]]"
        Dim string_ma83f As String = "PROJCS[" & Chr(34) & "MA83F" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Lambert_Conformal_Conic_2SP" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",656166.667],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",2460625.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-71.50000000000000],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",41.00000000000000],PARAMETER[" & Chr(34) & "standard_parallel_1" & Chr(34) & ",42.68333333333333],PARAMETER[" & Chr(34) & "standard_parallel_2" & Chr(34) & ",41.71666666666666],UNIT[" & Chr(34) & "Foot_US" & Chr(34) & ",0.30480060960122]]"
        Dim string_utm83_17f As String = "PROJCS[" & Chr(34) & "UTM83-17F" & Chr(34) & ",GEOGCS[" & Chr(34) & "LL83" & Chr(34) & ",DATUM[" & Chr(34) & "NAD83" & Chr(34) & ",SPHEROID[" & Chr(34) & "GRS1980" & Chr(34) & ",6378137.000,298.25722210]],PRIMEM[" & Chr(34) & "Greenwich" & Chr(34) & ",0],UNIT[" & Chr(34) & "Degree" & Chr(34) & ",0.017453292519943295]],PROJECTION[" & Chr(34) & "Transverse_Mercator" & Chr(34) & "],PARAMETER[" & Chr(34) & "false_easting" & Chr(34) & ",1640416.667],PARAMETER[" & Chr(34) & "false_northing" & Chr(34) & ",0.000],PARAMETER[" & Chr(34) & "central_meridian" & Chr(34) & ",-81.00000000000000],PARAMETER[" & Chr(34) & "scale_factor" & Chr(34) & ",0.9996],PARAMETER[" & Chr(34) & "latitude_of_origin" & Chr(34) & ",0.000],UNIT[" & Chr(34) & "Foot_US" & Chr(34) & ",0.30480060960122]]"



        Dim Coord_factory1 As New OSGeo.MapGuide.MgCoordinateSystemFactory
        Dim CoordSys1 As OSGeo.MapGuide.MgCoordinateSystem '= Coord_factory1.Create(Curent_system)
        Dim CoordSys2 As OSGeo.MapGuide.MgCoordinateSystem ' = Coord_factory1.Create(String_LL84)

        If RadioButton_10TM_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_CANA83_10TM115)
        If RadioButton_10tm_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_CANA83_10TM115)

        If RadioButton_3TM_111_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_CANA83_3TM111)
        If RadioButton_3TM_111_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_CANA83_3TM111)

        If RadioButton_3TM_114_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_CANA83_3TM114)
        If RadioButton_3TM_114_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_CANA83_3TM114)

        If RadioButton_3TM_117_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_CANA83_3TM117)
        If RadioButton_3TM_117_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_CANA83_3TM117)

        If RadioButton_3TM_120_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_CANA83_3TM120)
        If RadioButton_3TM_120_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_CANA83_3TM120)

        If RadioButton_UTM_83_11_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_UTM83_11)
        If RadioButton_utm_83_11_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_UTM83_11)

        If RadioButton_UTM_83_12_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_UTM83_12)
        If RadioButton_utm_83_12_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_UTM83_12)

        If RadioButton_LL84_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_LL84)
        If RadioButton_LL84_to.Checked = True Then CoordSys2 = Coord_factory1.Create(String_LL84)

        If RadioButton_UTM_27_11_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_UTM27_11)
        If RadioButton_UTM_27_11_TO.Checked = True Then CoordSys2 = Coord_factory1.Create(String_UTM27_11)

        If RadioButton_UTM_27_12_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_UTM27_12)
        If RadioButton_UTM_27_12_TO.Checked = True Then CoordSys2 = Coord_factory1.Create(String_UTM27_12)

        If RadioButton_UTM_83_10_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_UTM83_10)
        If RadioButton_UTM_83_10_TO.Checked = True Then CoordSys2 = Coord_factory1.Create(String_UTM83_10)

        If RadioButton_UTM_27_10_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_UTM27_10)
        If RadioButton_UTM_27_10_TO.Checked = True Then CoordSys2 = Coord_factory1.Create(String_UTM27_10)

        If RadioButton_UTM_83_9_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(String_UTM83_9)
        If RadioButton_UTM_83_9_TO.Checked = True Then CoordSys2 = Coord_factory1.Create(String_UTM83_9)

        If RadioButton_MA83F_FROM.Checked = True Then CoordSys1 = Coord_factory1.Create(string_ma83f)
        If RadioButton_MA83F_TO.Checked = True Then CoordSys2 = Coord_factory1.Create(string_ma83f)

        If RadioButton_utm_83_17F_from.Checked = True Then CoordSys1 = Coord_factory1.Create(string_utm83_17f)
        If RadioButton_utm_83_17F_to.Checked = True Then CoordSys2 = Coord_factory1.Create(string_utm83_17f)


        Dim Transform1 As OSGeo.MapGuide.MgCoordinateSystemTransform = Coord_factory1.GetTransform(CoordSys1, CoordSys2)
        Dim Coord1 As OSGeo.MapGuide.MgCoordinate '= Transform1.Transform(X5, Y5)

        Dim IsLL As Boolean = False
        If CoordSys2.Projection.ToString = "LL" Then
            IsLL = True
        End If

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet
        Try
            W1 = Get_active_worksheet_from_Excel()
            Dim Start1 As Double = Val(TextBox_start.Text)
            Dim End1 As Double = Val(TextBox_end.Text)

            If Start1 = 0 Then
                MsgBox("Please set the start row", , "Dan says...")
                Button_convert.Visible = True
                Exit Sub
            End If
            If End1 = 0 Then
                MsgBox("Please set the end row", , "Dan says...")
                Button_convert.Visible = True
                Exit Sub
            End If
            If End1 < Start1 Then
                MsgBox("End row smaller than start row", , "Dan says...")
                Button_convert.Visible = True
                Exit Sub
            End If
            Dim Column_x_from As String = TextBox_east_from.Text.ToUpper
            Dim Column_y_from As String = TextBox_NORTH_FROM.Text.ToUpper
            Dim Column_x_to As String = TextBox_Easting_to.Text.ToUpper
            Dim Column_y_to As String = TextBox_norting_to.Text.ToUpper
            Dim Problem_text As String = "You have problems on the following rows:"
            If Not Column_x_from = "" And Not Column_x_to = "" And Not Column_y_from = "" And Not Column_y_to = "" Then
                For i = Start1 To End1
                    Dim x As String = W1.Range(Column_x_from & i).Value
                    Dim y As String = W1.Range(Column_y_from & i).Value
                    If String.IsNullOrEmpty(x) = False Or String.IsNullOrEmpty(y) = False Then
                        If IsNumeric(x) = True And IsNumeric(y) = True Then
                            Coord1 = Transform1.Transform(Val(x), Val(y))
                            If IsLL = False Then
                                W1.Range(Column_x_to & i).Value = Math.Round(Coord1.X, 3)
                                W1.Range(Column_y_to & i).Value = Math.Round(Coord1.Y, 3)
                            Else
                                W1.Range(Column_x_to & i).Value = Math.Round(Coord1.X, 6)
                                W1.Range(Column_y_to & i).Value = Math.Round(Coord1.Y, 6)
                            End If

                        Else
                            W1.Range(Column_x_from & i).Interior.Color = 255
                            W1.Range(Column_y_from & i).Interior.Color = 255
                            Problem_text = Problem_text & vbCrLf & i
                        End If
                    Else
                        W1.Range(Column_x_from & i).Interior.Color = 255
                        W1.Range(Column_y_from & i).Interior.Color = 255
                        Problem_text = Problem_text & vbCrLf & i
                    End If

                Next
            End If
            If Problem_text <> "You have problems on the following rows:" Then
                MsgBox(Problem_text, , "Dan says...")
            End If
            Button_convert.Visible = True

        Catch ex As Exception
            Button_convert.Visible = True
            MsgBox(ex.Message)

        End Try


    End Sub



    Private Sub TextBox_start_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_start.KeyDown
        If e.KeyValue = 13 Then
            With TextBox_end
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub


    Private Sub TextBox_end_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_end.KeyDown
        If e.KeyValue = 13 Then
            Button_convert_Click(sender, e)
        End If
    End Sub

    Private Sub TextBox_east_from_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_east_from.KeyDown
        If e.KeyValue = 13 Then
            With TextBox_NORTH_FROM
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_NORTH_FROM_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_NORTH_FROM.KeyDown
        If e.KeyValue = 13 Then
            With TextBox_Easting_to
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_Easting_to_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_Easting_to.KeyDown
        If e.KeyValue = 13 Then
            With TextBox_norting_to
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub TextBox_norting_to_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_norting_to.KeyDown
        If e.KeyValue = 13 Then
            With TextBox_start
                .SelectAll()
                .Focus()
            End With
        End If
    End Sub

    Private Sub Button_Info_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Info.Click
        Dim String1 As String
        String1 = "This tool uses the internal engine of Autocad for conversions between the coordinate systems." & _
        vbCrLf & vbCrLf & "If your Autocad is not configured properly the result may be wrong." & _
        vbCrLf & "A known problem with Autocad is converting to/from UTM27. By default the conversion is turned off." & _
        vbCrLf & "In order for the conversion to work you need to enable the NAD 27 – NAD83 conversion." & _
        vbCrLf & vbCrLf & "Don’t use this tool to convert back and forth. The results are displayed rounded so an error due to this will occur." & _
        vbCrLf & vbCrLf & "This tool doesn’t use any information from the current drawing. "
        MsgBox(String1)

    End Sub
End Class