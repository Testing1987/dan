Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class SURVEY_TOOLS_FORM
    Dim dt1 As System.Data.DataTable
    Dim dt0 As System.Data.DataTable


    Dim dj_pipe As String = "dj_pipe"
    Dim heat1 As String = "heat1"
    Dim heat2 As String = "heat2"
    Dim serial_id As String = "serial_id"

    Dim dj_pipe0 As String = "pipe_no"
    Dim heat0 As String = "heat_no"
    Dim serial_id0 As String = "serial_no"
    Dim Found As String = "found_in_data"
    Dim mjoint As String = "multiple_joint_no"

    Dim data_table_scan As System.Data.DataTable
    Dim dt_comp As System.Data.DataTable

    Dim sta As String = "Station"
    Dim Descr As String = "Description"
    Dim Descr_scan As String = "Description from scan"
    Private Sub Button_reset_Click(sender As Object, e As EventArgs) Handles Button_reset.Click
        dt1 = New System.Data.DataTable
        dt1.Columns.Add(dj_pipe, GetType(String))
        dt1.Columns.Add(heat1, GetType(String))
        dt1.Columns.Add(heat2, GetType(String))
        dt1.Columns.Add(serial_id, GetType(String))
    End Sub

    Private Sub Button_reset2_Click(sender As Object, e As EventArgs) Handles Button_reset2.Click
        data_table_scan = New System.Data.DataTable
        data_table_scan.Columns.Add(sta, GetType(Double))
        data_table_scan.Columns.Add(Descr, GetType(String))
    End Sub
    Private Sub Button_load_data_Click(sender As Object, e As EventArgs) Handles Button_load_data.Click
        Try
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_start.Text)
            Dim end1 As Integer = CInt(TextBox_end.Text)

            If IsNothing(dt1) = True Then
                dt1 = New System.Data.DataTable
                dt1.Columns.Add(dj_pipe, GetType(String))
                dt1.Columns.Add(heat1, GetType(String))
                dt1.Columns.Add(heat2, GetType(String))
                dt1.Columns.Add(serial_id, GetType(String))
            End If

            Dim first_index As Integer = dt1.Rows.Count

            For i = start1 To end1
                dt1.Rows.Add()
            Next


            If Not TextBox_load_serial_pipe_id.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_load_serial_pipe_id.Text & start1 & ":" & TextBox_load_serial_pipe_id.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2
                Dim j As Integer = 0

                For i = first_index To dt1.Rows.Count - 1
                    If IsNothing(Values1(j + 1, 1)) = False Then
                        Dim String1 As String = Values1(j + 1, 1).ToString
                        If Not String1.ToLower = "n/a" Then
                            dt1.Rows(i).Item(serial_id) = String1
                        End If
                    End If
                    j = j + 1
                Next

            End If


            If Not TextBox_load_joint_pipe.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_load_joint_pipe.Text & start1 & ":" & TextBox_load_joint_pipe.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2
                Dim j As Integer = 0

                For i = first_index To dt1.Rows.Count - 1
                    If IsNothing(Values1(j + 1, 1)) = False Then
                        dt1.Rows(i).Item(dj_pipe) = Values1(j + 1, 1).ToString
                    End If
                    j = j + 1
                Next

            End If

            If Not TextBox_load_heat1.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_load_heat1.Text & start1 & ":" & TextBox_load_heat1.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2
                Dim j As Integer = 0
                Dim Last_index = dt1.Rows.Count - 1

                Dim i As Integer = first_index

                Do While i <= Last_index
                    If IsNothing(Values1(j + 1, 1)) = False Then
                        Dim String1 As String = Values1(j + 1, 1).ToString
                        If Not String1.ToLower = "n/a" Then

                            If String1.Contains("/") = False Then
                                dt1.Rows(i).Item(heat1) = String1
                            End If

                            If String1.Contains("/") = True Then
                                Dim Str() As String = String1.Split("/")
                                dt1.Rows(i).Item(heat1) = Str(0)

                                Dim k As Integer = 1
                                For k = 1 To Str.Length - 1
                                    Dim Row1 As System.Data.DataRow = dt1.NewRow
                                    If IsDBNull(dt1.Rows(i).Item(dj_pipe)) = False Then
                                        Row1.Item(dj_pipe) = dt1.Rows(i).Item(dj_pipe)
                                    End If

                                    If IsDBNull(dt1.Rows(i).Item(serial_id)) = False Then
                                        Row1.Item(serial_id) = dt1.Rows(i).Item(serial_id)
                                    End If
                                    Row1.Item(heat1) = Str(k)
                                    dt1.Rows.InsertAt(Row1, i + k)
                                Next
                                Last_index = Last_index + k - 1
                                i = i + k - 1
                            End If
                        End If
                    End If
                    i = i + 1
                    j = j + 1
                Loop
            End If

            If Not TextBox_load_heat2.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_load_heat2.Text & start1 & ":" & TextBox_load_heat2.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2
                Dim j As Integer = 0
                Dim Last_index = dt1.Rows.Count - 1

                Dim i As Integer = first_index

                Do While i <= Last_index
                    If IsNothing(Values1(j + 1, 1)) = False Then
                        Dim String1 As String = Values1(j + 1, 1).ToString
                        If Not String1.ToLower = "n/a" Then

                            If String1.Contains("/") = False Then
                                dt1.Rows(i).Item(heat2) = String1
                            End If

                            If String1.Contains("/") = True Then
                                Dim Str() As String = String1.Split("/")
                                dt1.Rows(i).Item(heat2) = Str(0)

                                Dim k As Integer = 1
                                For k = 1 To Str.Length - 1
                                    Dim Row1 As System.Data.DataRow = dt1.NewRow
                                    If IsDBNull(dt1.Rows(i).Item(dj_pipe)) = False Then
                                        Row1.Item(dj_pipe) = dt1.Rows(i).Item(dj_pipe)
                                    End If

                                    If IsDBNull(dt1.Rows(i).Item(serial_id)) = False Then
                                        Row1.Item(serial_id) = dt1.Rows(i).Item(serial_id)
                                    End If
                                    If IsDBNull(dt1.Rows(i).Item(heat1)) = False Then
                                        Row1.Item(heat1) = dt1.Rows(i).Item(heat1)
                                    End If

                                    Row1.Item(heat2) = Str(k)
                                    dt1.Rows.InsertAt(Row1, i + k)
                                Next
                                Last_index = Last_index + k - 1
                                i = i + k - 1
                            End If
                        End If
                    End If
                    i = i + 1
                    j = j + 1
                Loop
            End If

            Transfer_datatable_to_new_excel_spreadsheet(dt1)

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button_load_scan_Click(sender As Object, e As EventArgs) Handles Button_load_scan.Click
        Try
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_scan_startr.Text)
            Dim end1 As Integer = CInt(TextBox_scan_endr.Text)

            If IsNothing(data_table_scan) = True Then
                data_table_scan = New System.Data.DataTable
                data_table_scan.Columns.Add(sta, GetType(Double))
                data_table_scan.Columns.Add(Descr, GetType(String))
            End If

            Dim first_index As Integer = data_table_scan.Rows.Count

            For i = start1 To end1
                data_table_scan.Rows.Add()
            Next



            Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_scan_sta.Text & start1 & ":" & TextBox_scan_sta.Text & end1)
            Dim Values1(end1 - start1, 0) As Object
            Values1 = Range1.Value2
            Dim j As Integer = 0

            For i = first_index To data_table_scan.Rows.Count - 1
                If IsNothing(Values1(j + 1, 1)) = False Then
                    Dim String1 As String = Values1(j + 1, 1).ToString
                    If IsNumeric(String1) = True Then
                        data_table_scan.Rows(i).Item(sta) = CDbl(String1)
                    End If
                End If
                j = j + 1
            Next





            Range1 = W1.Range(TextBox_scan_layer.Text & start1 & ":" & TextBox_scan_layer.Text & end1)
            Dim Values2(end1 - start1, 0) As Object
            Values2 = Range1.Value2
            j = 0

            For i = first_index To data_table_scan.Rows.Count - 1
                If IsNothing(Values2(j + 1, 1)) = False Then
                    data_table_scan.Rows(i).Item(Descr) = Values2(j + 1, 1).ToString
                End If
                j = j + 1
            Next





            Transfer_datatable_to_new_excel_spreadsheet(data_table_scan)

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_compare_Click(sender As Object, e As EventArgs) Handles Button_compare.Click
        Try
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_start.Text)
            Dim end1 As Integer = CInt(TextBox_end.Text)


            dt0 = New System.Data.DataTable

            dt0.Columns.Add(Found, GetType(String))
            dt0.Columns.Add(mjoint, GetType(String))
            dt0.Columns.Add(dj_pipe0, GetType(String))
            dt0.Columns.Add(serial_id0, GetType(String))
            dt0.Columns.Add(heat0, GetType(String))



            For i = start1 To end1
                dt0.Rows.Add()
                dt0.Rows(dt0.Rows.Count - 1).Item(Found) = "NO"
            Next


            If Not TextBox_match_serial_no.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_match_serial_no.Text & start1 & ":" & TextBox_match_serial_no.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2

                For i = 0 To dt0.Rows.Count - 1
                    If IsNothing(Values1(i + 1, 1)) = False Then
                        Dim String1 As String = Values1(i + 1, 1).ToString
                        dt0.Rows(i).Item(serial_id0) = String1
                    End If
                Next
            End If

            If Not TextBox_match_heat_no.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_match_heat_no.Text & start1 & ":" & TextBox_match_heat_no.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2

                For i = 0 To dt0.Rows.Count - 1
                    If IsNothing(Values1(i + 1, 1)) = False Then
                        Dim String1 As String = Values1(i + 1, 1).ToString
                        dt0.Rows(i).Item(heat0) = String1
                    End If

                Next
            End If

            If Not TextBox_match_joint_no.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_match_joint_no.Text & start1 & ":" & TextBox_match_joint_no.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2

                For i = 0 To dt0.Rows.Count - 1
                    If IsNothing(Values1(i + 1, 1)) = False Then
                        Dim String1 As String = Values1(i + 1, 1).ToString
                        dt0.Rows(i).Item(mjoint) = String1
                    End If

                Next
            End If

            If Not TextBox_match_pipe_no.Text = "" Then
                Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_match_pipe_no.Text & start1 & ":" & TextBox_match_pipe_no.Text & end1)
                Dim Values1(end1 - start1, 0) As Object
                Values1 = Range1.Value2

                For i = 0 To dt0.Rows.Count - 1
                    If IsNothing(Values1(i + 1, 1)) = False Then
                        Dim String1 As String = Values1(i + 1, 1).ToString
                        dt0.Rows(i).Item(dj_pipe0) = String1
                    End If

                Next
            End If

            If IsNothing(dt1) = True Then
                MsgBox("No data table for comparation was loaded")
                Exit Sub
            Else
                If dt1.Rows.Count = 0 Then
                    MsgBox("No data table for comparation was loaded")
                    Exit Sub
                End If
            End If


            For i = 0 To dt0.Rows.Count - 1
                Dim Mjoint0 As String = ""
                If IsDBNull(dt0.Rows(i).Item(mjoint)) = False Then
                    Mjoint0 = dt0.Rows(i).Item(mjoint)
                End If
                Dim Pipe0 As String = ""
                If IsDBNull(dt0.Rows(i).Item(dj_pipe0)) = False Then
                    Pipe0 = dt0.Rows(i).Item(dj_pipe0)
                End If
                Dim serial0 As String = ""
                If IsDBNull(dt0.Rows(i).Item(serial_id0)) = False Then
                    serial0 = dt0.Rows(i).Item(serial_id0)
                End If
                Dim ht0 As String = ""
                If IsDBNull(dt0.Rows(i).Item(heat0)) = False Then
                    ht0 = dt0.Rows(i).Item(heat0)
                End If

                For j = 0 To dt1.Rows.Count - 1
                    Dim dj1 As String = ""
                    If IsDBNull(dt1.Rows(j).Item(dj_pipe)) = False Then
                        dj1 = dt1.Rows(j).Item(dj_pipe)
                    End If
                    Dim ht1 As String = ""
                    If IsDBNull(dt1.Rows(j).Item(heat1)) = False Then
                        ht1 = dt1.Rows(j).Item(heat1)
                    End If
                    Dim ht2 As String = ""
                    If IsDBNull(dt1.Rows(j).Item(heat2)) = False Then
                        ht2 = dt1.Rows(j).Item(heat2)
                    End If
                    Dim serial1 As String = ""
                    If IsDBNull(dt1.Rows(j).Item(serial_id)) = False Then
                        serial1 = dt1.Rows(j).Item(serial_id)
                    End If

                    If serial0 = serial1 And Not serial0 = "" Then
                        dt0.Rows(i).Item(Found) = "YES"
                        dt0.Rows(i).Item(mjoint) = dj1
                        dt1.Rows(j).Delete()
                        Exit For
                    Else
                        If Pipe0 = dj1 And (ht0 = ht1 Or ht0 = ht2) And Not Pipe0 = "" And Not ht0 = "" Then
                            dt0.Rows(i).Item(Found) = "YES"
                            dt1.Rows(j).Delete()
                            Exit For
                        End If
                    End If


                Next

            Next

            Transfer_datatable_to_new_excel_spreadsheet(dt0)
            Transfer_datatable_to_new_excel_spreadsheet(dt1)

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_compare_scan_Click(sender As Object, e As EventArgs) Handles Button_compare_data_scan.Click
        Try
            Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()
            Dim start1 As Integer = CInt(TextBox_scan_startr.Text)
            Dim end1 As Integer = CInt(TextBox_scan_endr.Text)


            dt_comp = New System.Data.DataTable

            dt_comp.Columns.Add(sta, GetType(Double))
            dt_comp.Columns.Add(Descr, GetType(String))
            dt_comp.Columns.Add(Descr_scan, GetType(String))
            dt_comp.Columns.Add(Found, GetType(String))


            For i = start1 To end1
                dt_comp.Rows.Add()
                dt_comp.Rows(dt_comp.Rows.Count - 1).Item(Found) = "NO"
            Next


            Dim Range1 As Microsoft.Office.Interop.Excel.Range = W1.Range(TextBox_comp_sta.Text & start1 & ":" & TextBox_comp_sta.Text & end1)
            Dim Values1(end1 - start1, 0) As Object
            Values1 = Range1.Value2

            For i = 0 To dt_comp.Rows.Count - 1
                If IsNothing(Values1(i + 1, 1)) = False Then
                    Dim String1 As String = Values1(i + 1, 1).ToString
                    If IsNumeric(String1) = True Then
                        dt_comp.Rows(i).Item(sta) = CDbl(String1)
                    End If
                End If
            Next

            Range1 = W1.Range(TextBox_comp_descr.Text & start1 & ":" & TextBox_comp_descr.Text & end1)
            Dim Values2(end1 - start1, 0) As Object
            Values2 = Range1.Value2

            For i = 0 To dt_comp.Rows.Count - 1
                If IsNothing(Values2(i + 1, 1)) = False Then
                    Dim String1 As String = Values2(i + 1, 1).ToString

                    dt_comp.Rows(i).Item(Descr) = String1

                End If
            Next
           

            If IsNothing(data_table_scan) = True Then
                MsgBox("No data table for comparation was loaded")
                Exit Sub
            Else
                If data_table_scan.Rows.Count = 0 Then
                    MsgBox("No data table for comparation was loaded")
                    Exit Sub
                End If
            End If


            For i = 0 To dt_comp.Rows.Count - 1
                Dim Station1 As Double = 0
                If IsDBNull(dt_comp.Rows(i).Item(sta)) = False Then
                    Station1 = dt_comp.Rows(i).Item(sta)


                    For j = 0 To data_table_scan.Rows.Count - 1
                        Dim station2 As Double = 0
                        If IsDBNull(data_table_scan.Rows(j).Item(sta)) = False Then
                            station2 = data_table_scan.Rows(j).Item(sta)
                        End If
                        Dim ht1 As String = ""


                        If Round(Station1, 0) = Round(station2, 0) Then
                            dt_comp.Rows(i).Item(Found) = "YES"
                            dt_comp.Rows(i).Item(Descr_scan) = data_table_scan.Rows(j).Item(Descr)
                            data_table_scan.Rows(j).Delete()
                            Exit For
                        End If


                    Next
                End If

            Next

            Transfer_datatable_to_new_excel_spreadsheet(dt_comp)
            Transfer_datatable_to_new_excel_spreadsheet(data_table_scan)

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class