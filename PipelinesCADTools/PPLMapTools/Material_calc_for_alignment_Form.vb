Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Public Class Material_calc_for_alignment_Form
    Dim Colectie1 As New Specialized.StringCollection
    Dim Data_table1 As System.Data.DataTable
    Dim Data_table2 As System.Data.DataTable
    Dim Data_table3 As System.Data.DataTable
    Dim Index_data_table As Integer = 0
    Dim Index_data_table2 As Integer = 0
    Dim Index_data_table3 As Integer = 0

    Dim Index_row_excel As Integer
    Private Sub Material_calc_for_alignment_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Data_table1 = New System.Data.DataTable
        Data_table1.Columns.Add("TYPE", GetType(Double))
        Data_table1.Columns.Add("LENGTH", GetType(Double))
        Data_table1.Columns.Add("TRUE_LENGTH", GetType(Double))
        Data_table2 = New System.Data.DataTable
        Data_table2.Columns.Add("TYPE", GetType(Double))
        Data_table2.Columns.Add("LENGTH", GetType(Double))
        Data_table3 = New System.Data.DataTable
        Data_table3.Columns.Add("TYPE", GetType(Double))
        Data_table3.Columns.Add("LENGTH", GetType(Double))

        Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
    End Sub


    Private Sub Button_pick_Click(sender As Object, e As EventArgs) Handles Button_pick.Click

        Dim Empty_array() As ObjectId

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        Using lock As DocumentLock = ThisDrawing.LockDocument
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
1234:

                    Dim Rezultat_mat As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select Material Type:"

                    Object_Prompt.SingleOnly = False

                    Rezultat_mat = Editor1.GetSelection(Object_Prompt)



                    Dim Rezultat_chain_dist As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt2 As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt2.MessageForAdding = vbLf & "Select 2 Chainages or 1 distance:"

                    Object_Prompt2.SingleOnly = False

                    Rezultat_chain_dist = Editor1.GetSelection(Object_Prompt2)


                    If Rezultat_chain_dist.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If



                    If Rezultat_chain_dist.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        If IsNothing(Rezultat_chain_dist) = False Then

                            Dim Continut_mat As Integer = 99
                            Dim Continut_length As Double = -1
                            Dim Station1 As String = "xxx"
                            Dim Station2 As String = "xxx"

                            If Rezultat_mat.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat_mat) = False Then
                                    For i = 0 To Rezultat_mat.Value.Count - 1
                                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                        Obj1 = Rezultat_mat.Value.Item(i)
                                        Dim Ent1 As Entity
                                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                            Dim Text1 As DBText = Ent1
                                            If IsNumeric(Text1.TextString) = True Then
                                                Continut_mat = CInt(Text1.TextString)

                                                Exit For
                                            End If
                                        End If

                                        If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                            Dim mText1 As MText = Ent1
                                            If IsNumeric(mText1.Text) = True Then
                                                Continut_mat = CInt(mText1.Text)

                                                Exit For
                                            End If
                                        End If

                                        If TypeOf Ent1 Is BlockReference Then
                                            Dim Block1 As BlockReference = Ent1
                                            If Block1.AttributeCollection.Count > 0 Then
                                                Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                                For Each id In attColl
                                                    Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                    If attref.Tag.ToUpper = "MAT" Then
                                                        Dim Continut As String = attref.TextString
                                                        If IsNumeric(Continut) = True Then
                                                            Continut_mat = CInt(Continut)

                                                            Exit For
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If

                                    Next
                                End If
                            End If


                            If Rezultat_chain_dist.Value.Count = 1 Then
                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat_chain_dist.Value.Item(0)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Dim Text1 As DBText = Ent1
                                    If IsNumeric(Text1.TextString) = True Then
                                        Continut_length = CDbl(Text1.TextString)
                                    End If
                                End If

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                    Dim mText1 As MText = Ent1
                                    If IsNumeric(mText1.Text) = True Then
                                        Continut_length = CDbl(mText1.Text)
                                    End If
                                End If

                                If TypeOf Ent1 Is BlockReference Then
                                    Dim Block1 As BlockReference = Ent1
                                    If Block1.AttributeCollection.Count > 0 Then
                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                        Dim Begin_number, End_number As Double
                                        Dim Begin_string, End_string As String

                                        For Each id In attColl
                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                            Dim Continut As String = attref.TextString
                                            If attref.Tag.ToUpper = "LENGTH" Then
                                                If IsNumeric(Continut) = True Then
                                                    Continut_length = CDbl(Continut)
                                                End If
                                            End If
                                            If attref.Tag.ToUpper = "BEGINSTA" Then
                                                Begin_string = Continut
                                                Station1 = Continut
                                            End If
                                            If attref.Tag.ToUpper = "ENDSTA" Then
                                                End_string = Continut
                                                Station2 = Continut
                                            End If
                                        Next
                                        Dim stars As String = ""

                                        If IsNumeric(Replace(Begin_string, "+", "")) = True Then
                                            If IsNumeric(Replace(End_string, "+", "")) = True Then
                                                Begin_number = CDbl(Replace(Begin_string, "+", ""))
                                                End_number = CDbl(Replace(End_string, "+", ""))
                                                If Not Continut_length = -1 Then
                                                    If Not Continut_length = Round(Abs(Begin_number - End_number), 1) Then
                                                        MsgBox("ERROR in Length Station" & vbCrLf & Begin_string & "   ---   " & End_string)
                                                    End If
                                                End If

                                                Continut_length = Round(Abs(Begin_number - End_number), 1)


                                            End If
                                        End If

                                    End If
                                End If


                                If TypeOf Ent1 Is AttributeDefinition Then
                                    Dim attref As AttributeDefinition = Ent1
                                    Dim Continut As String = attref.Tag.ToString
                                    If IsNumeric(Continut) = True Then
                                        Continut_length = CDbl(Continut)
                                    End If
                                End If
                            End If




                            If Rezultat_chain_dist.Value.Count > 1 Then
                                Dim Start_string As String = ""
                                Dim End_string As String = ""
                                Dim Start_string_blocks1 As String = ""
                                Dim End_string_blocks1 As String = ""
                                Dim STATION_string_blocks1 As String = ""
                                Dim Length_blocks1 As Double = -1
                                Dim Start_string_blocks2 As String = ""
                                Dim End_string_blocks2 As String = ""
                                Dim STATION_string_blocks2 As String = ""
                                Dim Nr_blocks As Integer = 0
                                Dim Length_blocks2 As Double = -1
                                Dim XB1 As Double = -1
                                Dim xB2 As Double = -1
                                Dim xt As Double = -1
                                Dim xAT1 As Double = -1
                                Dim xAT11 As Double = -1
                                Dim xAT2 As Double = -1
                                Dim xAT22 As Double = -1

                                For i = 0 To Rezultat_chain_dist.Value.Count - 1
                                    Dim Obj2 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj2 = Rezultat_chain_dist.Value.Item(i)
                                    Dim Ent2 As Entity
                                    Ent2 = Obj2.ObjectId.GetObject(OpenMode.ForRead)

                                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                        Dim Text1 As DBText = Ent2
                                        xt = Text1.Position.X
                                        If Start_string = "" Then
                                            Start_string = Text1.TextString

                                        Else
                                            End_string = Text1.TextString
                                            Exit For
                                        End If

                                    End If

                                    If TypeOf Ent2 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                        Dim mText1 As MText = Ent2
                                        xt = mText1.Location.X
                                        If Start_string = "" Then
                                            Start_string = mText1.Text
                                        Else
                                            End_string = mText1.Text
                                            Exit For
                                        End If
                                    End If

                                    If TypeOf Ent2 Is BlockReference Then
                                        Dim Block1 As BlockReference = Ent2
                                        If Block1.AttributeCollection.Count > 0 Then
                                            If XB1 = -1 Then
                                                XB1 = Block1.Position.X
                                            Else
                                                xB2 = Block1.Position.X
                                            End If
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection


                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                Dim Continut As String = attref.TextString
                                                If Not Replace(Continut, " ", "") = "" Then
                                                    If attref.Tag.ToUpper = "LENGTH" Then
                                                        If IsNumeric(Continut) = True Then
                                                            If Nr_blocks = 0 Then
                                                                If Length_blocks1 = -1 Then
                                                                    Length_blocks1 = Round(CDbl(Continut), 1)
                                                                Else
                                                                    Length_blocks2 = Round(CDbl(Continut), 1)
                                                                End If
                                                            Else
                                                                Length_blocks2 = Round(CDbl(Continut), 1)
                                                            End If
                                                        End If
                                                    End If

                                                    If attref.Tag.ToUpper = "BEGINSTA" Then
                                                        If Nr_blocks = 0 Then
                                                            If Start_string_blocks1 = "" Then
                                                                Start_string_blocks1 = Continut
                                                            Else
                                                                Start_string_blocks2 = Continut
                                                            End If
                                                            If xAT1 = -1 Then
                                                                xAT1 = attref.Position.X
                                                            Else
                                                                xAT2 = attref.Position.X
                                                            End If
                                                        Else
                                                            Start_string_blocks2 = Continut
                                                            xAT2 = attref.Position.X
                                                        End If


                                                    End If
                                                    If attref.Tag.ToUpper = "ENDSTA" Then
                                                        If Nr_blocks = 0 Then
                                                            If End_string_blocks1 = "" Then
                                                                End_string_blocks1 = Continut
                                                            Else
                                                                End_string_blocks2 = Continut
                                                            End If
                                                            If xAT11 = -1 Then
                                                                xAT11 = attref.Position.X
                                                            Else
                                                                xAT22 = attref.Position.X
                                                            End If
                                                        Else
                                                            End_string_blocks2 = Continut
                                                            xAT22 = attref.Position.X
                                                        End If
                                                    End If

                                                    If attref.Tag.ToUpper = "STA" Then
                                                        If Nr_blocks = 0 Then
                                                            STATION_string_blocks1 = Continut
                                                        Else
                                                            STATION_string_blocks2 = Continut
                                                        End If

                                                    End If
                                                End If
                                            Next
                                            Nr_blocks = Nr_blocks + 1
                                        End If
                                    End If


                                    If TypeOf Ent2 Is AttributeDefinition Then
                                        Dim attref As AttributeDefinition = Ent2
                                        Dim Continut As String = attref.Tag.ToString
                                        If IsNumeric(Continut) = True Then
                                            Continut_length = CDbl(Continut)
                                            Exit For
                                        End If
                                    End If
                                Next

                                If Continut_length = -1 Then
                                    If Start_string = "" Then
                                        If Not Start_string_blocks1 = "" And Not Start_string_blocks2 = "" And Not End_string_blocks1 = "" And Not End_string_blocks2 = "" Then
                                            If IsNumeric(Replace(Start_string_blocks1, "+", "")) = True And IsNumeric(Replace(Start_string_blocks2, "+", "")) = True And IsNumeric(Replace(End_string_blocks1, "+", "")) = True And IsNumeric(Replace(End_string_blocks2, "+", "")) = True Then
                                                Dim sTART1, sTART2, eND1, eND2 As Double
                                                sTART1 = Round(CDbl(Replace(Start_string_blocks1, "+", "")), 1)
                                                sTART2 = Round(CDbl(Replace(Start_string_blocks2, "+", "")), 1)
                                                eND1 = Round(CDbl(Replace(End_string_blocks1, "+", "")), 1)
                                                eND2 = Round(CDbl(Replace(End_string_blocks2, "+", "")), 1)

                                                If xAT1 > xAT11 Then
                                                    Dim tEMP As Double
                                                    tEMP = sTART1
                                                    sTART1 = eND1
                                                    eND1 = tEMP
                                                End If


                                                If xAT2 > xAT22 Then
                                                    Dim tEMP As Double
                                                    tEMP = sTART2
                                                    sTART2 = eND2
                                                    eND2 = tEMP
                                                End If


                                                If Not XB1 = -1 And Not xB2 = -1 Then
                                                    If XB1 < xB2 Then
                                                        Continut_length = Round(Abs(eND1 - sTART2), 1)
                                                        Station1 = Get_chainage_from_double(eND1, 1)
                                                        Station2 = Get_chainage_from_double(sTART2, 1)
                                                        Start_string = ""
                                                        End_string = ""
                                                    Else
                                                        Continut_length = Round(Abs(eND2 - sTART1), 1)
                                                        Station1 = Get_chainage_from_double(eND2, 1)
                                                        Station2 = Get_chainage_from_double(sTART1, 1)
                                                        Start_string = ""
                                                        End_string = ""
                                                    End If
                                                End If
                                                If Not Length_blocks1 = -1 Then
                                                    If Not Round(Length_blocks1, 1) = Round(Abs(sTART1 - eND1), 1) Then
                                                        MsgBox("ERROR in Length Station" & vbCrLf & Start_string_blocks1 & "   ---   " & End_string_blocks1)
                                                    End If
                                                End If
                                                If Not Length_blocks2 = -1 Then
                                                    If Not Round(Length_blocks2, 1) = Round(Abs(sTART2 - eND2), 1) Then
                                                        MsgBox("ERROR in Length Station" & vbCrLf & Start_string_blocks2 & "   ---   " & End_string_blocks2)
                                                    End If
                                                End If

                                            End If

                                        End If

                                        If (Not Start_string_blocks1 = "" Or Not End_string_blocks1 = "") And (Not Start_string_blocks2 = "" Or Not End_string_blocks2 = "") _
                                            And (Start_string_blocks1 = "" Or End_string_blocks1 = "" Or Start_string_blocks2 = "" Or End_string_blocks2 = "") Then

                                            Dim sTART1 As Double = -1
                                            Dim sTART2 As Double = -1
                                            Dim eND1 As Double = -1
                                            Dim eND2 As Double = -1

                                            If IsNumeric(Replace(Start_string_blocks1, "+", "")) = True Then
                                                sTART1 = Round(CDbl(Replace(Start_string_blocks1, "+", "")), 1)
                                            End If

                                            If IsNumeric(Replace(Start_string_blocks2, "+", "")) = True Then
                                                sTART2 = Round(CDbl(Replace(Start_string_blocks2, "+", "")), 1)
                                            End If
                                            If IsNumeric(Replace(End_string_blocks1, "+", "")) = True Then
                                                eND1 = Round(CDbl(Replace(End_string_blocks1, "+", "")), 1)
                                            End If
                                            If IsNumeric(Replace(End_string_blocks2, "+", "")) = True Then
                                                eND2 = Round(CDbl(Replace(End_string_blocks2, "+", "")), 1)
                                            End If

                                            If Not xAT1 = -1 And Not xAT11 = -1 Then
                                                If xAT1 > xAT11 Then
                                                    Dim tEMP As Double
                                                    tEMP = sTART1
                                                    sTART1 = eND1
                                                    eND1 = tEMP
                                                End If
                                            End If

                                            If Not xAT2 = -2 And Not xAT22 = -2 Then
                                                If xAT2 > xAT22 Then
                                                    Dim tEMP As Double
                                                    tEMP = sTART2
                                                    sTART2 = eND2
                                                    eND2 = tEMP
                                                End If
                                            End If

                                            If sTART1 = -1 And Not eND1 = -1 Then
                                                sTART1 = eND1
                                            End If
                                            If sTART2 = -1 And Not eND2 = -1 Then
                                                sTART2 = eND2
                                            End If
                                            If Not sTART1 = -1 And eND1 = -1 Then
                                                eND1 = sTART1
                                            End If
                                            If Not sTART2 = -1 And eND2 = -1 Then
                                                eND2 = sTART2
                                            End If

                                            If Not XB1 = -1 And Not xB2 = -1 Then
                                                If XB1 < xB2 Then
                                                    Continut_length = Round(Abs(eND1 - sTART2), 1)
                                                    Station1 = Get_chainage_from_double(eND1, 1)
                                                    Station2 = Get_chainage_from_double(sTART2, 1)
                                                    Start_string = ""
                                                    End_string = ""
                                                Else
                                                    Continut_length = Round(Abs(eND2 - sTART1), 1)
                                                    Station1 = Get_chainage_from_double(eND2, 1)
                                                    Station2 = Get_chainage_from_double(sTART1, 1)
                                                    Start_string = ""
                                                    End_string = ""
                                                End If
                                            End If
                                        End If

                                    End If

                                    If Not Start_string = "" And End_string = "" Then
                                        If Not Start_string_blocks1 = "" Or Not End_string_blocks1 = "" Or Not STATION_string_blocks1 = "" Then
                                            If Not XB1 = -1 And Not xt = -1 Then
                                                If xt < XB1 Then
                                                    If Not Start_string_blocks1 = "" Then
                                                        End_string = Start_string_blocks1
                                                    Else
                                                        If Not STATION_string_blocks1 = "" Then
                                                            End_string = STATION_string_blocks1
                                                        Else
                                                            If Not End_string_blocks1 = "" Then
                                                                End_string = End_string_blocks1
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If Not End_string_blocks1 = "" Then
                                                        End_string = End_string_blocks1
                                                    Else
                                                        If Not STATION_string_blocks1 = "" Then
                                                            End_string = STATION_string_blocks1
                                                        Else
                                                            If Not Start_string_blocks1 = "" Then
                                                                End_string = Start_string_blocks1
                                                            End If
                                                        End If
                                                    End If


                                                End If
                                            End If
                                        End If
                                    End If

                                    If Start_string = "" And Not End_string = "" Then
                                        If Not Start_string_blocks1 = "" Or Not End_string_blocks1 = "" Or Not STATION_string_blocks1 = "" Then
                                            If Not XB1 = -1 And Not xt = -1 Then
                                                If xt < XB1 Then
                                                    If Not Start_string_blocks1 = "" Then
                                                        Start_string = Start_string_blocks1
                                                    Else
                                                        If Not STATION_string_blocks1 = "" Then
                                                            Start_string = STATION_string_blocks1
                                                        Else
                                                            If Not End_string_blocks1 = "" Then
                                                                Start_string = End_string_blocks1
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If Not End_string_blocks1 = "" Then
                                                        Start_string = End_string_blocks1
                                                    Else
                                                        If Not STATION_string_blocks1 = "" Then
                                                            Start_string = STATION_string_blocks1
                                                        Else
                                                            If Not Start_string_blocks1 = "" Then
                                                                Start_string = Start_string_blocks1
                                                            End If
                                                        End If
                                                    End If


                                                End If
                                            End If
                                        End If
                                    End If

                                    If Not Start_string = "" And Not End_string = "" Then
                                        If IsNumeric(Replace(Start_string, "+", "")) = True And IsNumeric(Replace(End_string, "+", "")) = True Then
                                            Continut_length = Abs(Round(CDbl(Replace(Start_string, "+", "")) - CDbl(Replace(End_string, "+", "")), 1))
                                            Station1 = Start_string
                                            Station2 = End_string
                                        End If
                                    End If


                                End If

                            End If

                            If Not Continut_mat = 0 And Not Continut_length = -1 Then
                                ListBox_mat_type.Items.Add(Continut_mat)
                                ListBox_lenghts.Items.Add(Continut_length)
                                ListBox_chainage1.Items.Add(Station1)
                                ListBox_chainage2.Items.Add(Station2)
                                Data_table1.Rows.Add()
                                Data_table1.Rows(Index_data_table).Item("TYPE") = Continut_mat
                                Data_table1.Rows(Index_data_table).Item("LENGTH") = Continut_length
                                Index_data_table = Index_data_table + 1
                            End If


                        End If
                    End If



                    GoTo 1234
                End Using
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox(ex.Message)
            End Try
        End Using


    End Sub

    Private Sub Button_calc_totals_Click(sender As Object, e As EventArgs) Handles Button_calc_totals.Click
        If ListBox_mat_type.Items.Count > 0 Then

            ListBox_Totals.Items.Clear()

            Data_table2.Rows.Clear()
            Index_data_table2 = 0



            Dim groupedTable = From table In Data_table1 Group table By groupfield = table.Field(Of Double)("TYPE") Into myGroup = Group Select New With {Key groupfield, .TotalAmount = myGroup.Sum(Function(r) r.Field(Of Double)("LENGTH"))}
            Dim Total_gen As Double = 0
            For Each row In groupedTable

                ListBox_Totals.Items.Add(row.groupfield & " - " & Round(row.TotalAmount, 1).ToString)
                Data_table2.Rows.Add()
                Data_table2.Rows(Index_data_table2).Item("TYPE") = row.groupfield
                Data_table2.Rows(Index_data_table2).Item("LENGTH") = row.TotalAmount
                Index_data_table2 = Index_data_table2 + 1
                Total_gen = Total_gen + row.TotalAmount
            Next
            ListBox_Totals.Items.Add("___")
            ListBox_Totals.Items.Add(Round(Total_gen, 1))

            ListBox_true_length.Items.Clear()
            Data_table3.Rows.Clear()
            Index_data_table3 = 0
            Dim groupedTable3 = From table In Data_table1 Group table By groupfield = table.Field(Of Double)("TYPE") Into myGroup = Group Select New With {Key groupfield, .TotalAmount = myGroup.Sum(Function(r) r.Field(Of Double)("TRUE_LENGTH"))}
            Dim Total_gen3 As Double = 0
            For Each row3 In groupedTable3

                ListBox_true_length.Items.Add(row3.groupfield & " - " & Round(row3.TotalAmount, 1).ToString)
                Data_table3.Rows.Add()
                Data_table3.Rows(Index_data_table3).Item("TYPE") = row3.groupfield
                Data_table3.Rows(Index_data_table3).Item("LENGTH") = row3.TotalAmount
                Index_data_table3 = Index_data_table3 + 1
                Total_gen3 = Total_gen3 + row3.TotalAmount
            Next
            ListBox_true_length.Items.Add("___")
            ListBox_true_length.Items.Add(Round(Total_gen3, 1))


        End If
    End Sub



    Private Sub Button_clear_Click(sender As Object, e As EventArgs) Handles Button_clear.Click
        ListBox_lenghts.Items.Clear()
        ListBox_mat_type.Items.Clear()
        ListBox_chainage1.Items.Clear()
        ListBox_chainage2.Items.Clear()
        ListBox_Totals.Items.Clear()
        ListBox_true_length.Items.Clear()
        Data_table1.Rows.Clear()
        Index_data_table = 0
        Data_table3.Rows.Clear()
        Data_table2.Rows.Clear()
    End Sub

    Private Sub ListBox_mat_type_Click(sender As Object, e As EventArgs) Handles ListBox_mat_type.Click
        Try
            Dim curent_index As Integer = ListBox_mat_type.SelectedIndex
            If curent_index >= 0 Then
                If ListBox_mat_type.Items.Count > 0 Then
                    Dim Rezultat_msg As MsgBoxResult = MsgBox("Delete?", vbYesNo)
                    If Rezultat_msg = vbYes Then
                        ListBox_mat_type.Items.RemoveAt(curent_index)
                        ListBox_lenghts.Items.RemoveAt(curent_index)
                        ListBox_chainage1.Items.RemoveAt(curent_index)
                        ListBox_chainage2.Items.RemoveAt(curent_index)
                        Data_table1.Rows(curent_index).Delete()
                        Index_data_table = Index_data_table - 1
                        If Index_data_table = -1 Then Index_data_table = 0
                    Else
                        ListBox_chainage1.SelectedIndex = curent_index
                        ListBox_chainage2.SelectedIndex = curent_index
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub ListBox_lenghts_Click(sender As Object, e As EventArgs) Handles ListBox_lenghts.Click
        Try
            Dim curent_index As Integer = ListBox_lenghts.SelectedIndex
            If curent_index >= 0 Then
                If ListBox_lenghts.Items.Count > 0 Then

                    Dim Rezultat_msg As MsgBoxResult = MsgBox("Edit?", vbYesNo)
                    If Rezultat_msg = vbYes Then
                        Dim Len1 As String = InputBox("Specify new length:")
                        If Not Len1 = "" Then
                            If IsNumeric(Len1) = True Then
                                ListBox_lenghts.Items(curent_index) = Len1
                                Data_table1.Rows(curent_index).Item("LENGTH") = Len1
                            End If
                        End If
                    End If
                    ListBox_chainage1.SelectedIndex = curent_index
                    ListBox_chainage2.SelectedIndex = curent_index
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_TRANSFER_TOTALS_Click(sender As Object, e As EventArgs) Handles Button_TRANSFER_TOTALS.Click
        Dim Empty_array() As ObjectId

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        Using lock As DocumentLock = ThisDrawing.LockDocument
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction

                    For i = 0 To Data_table3.Rows.Count - 1
                        Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                        Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                        Object_Prompt.MessageForAdding = vbLf & "Select Material " & Data_table3.Rows(i).Item("TYPE")

                        Object_Prompt.SingleOnly = True

                        Rezultat1 = Editor1.GetSelection(Object_Prompt)
                        If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                            If IsNothing(Rezultat1) = False Then
                                For j = 0 To Rezultat1.Value.Count - 1
                                    Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                    Obj1 = Rezultat1.Value.Item(j)
                                    Dim Ent1 As Entity
                                    Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                        Dim Text1 As DBText = Ent1
                                        Text1.UpgradeOpen()
                                        Text1.TextString = Get_String_Rounded(Data_table3.Rows(i).Item("LENGTH"), 1)
                                    End If

                                    If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                                        Dim mText1 As MText = Ent1
                                        mText1.UpgradeOpen()
                                        mText1.Contents = Get_String_Rounded(Data_table3.Rows(i).Item("LENGTH"), 1)
                                    End If
                                Next
                            End If
                        End If


                    Next

                    Trans1.Commit()

                End Using
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox(ex.Message)
            End Try
        End Using

    End Sub

    Private Sub Button_pick_all_Click(sender As Object, e As EventArgs) Handles Button_pick_all.Click
        Dim Empty_array() As ObjectId

        Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor = ThisDrawing.Editor

        Using lock As DocumentLock = ThisDrawing.LockDocument
            ascunde_butoanele_pentru_forms(Me, Colectie1)
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Try


                Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
1234:

                    Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult

                    Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    Object_Prompt.MessageForAdding = vbLf & "Select Blocks:"

                    Object_Prompt.SingleOnly = False

                    Rezultat1 = Editor1.GetSelection(Object_Prompt)



                    If Rezultat1.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                        Editor1.WriteMessage(vbLf & "Command:")
                        afiseaza_butoanele_pentru_forms(Me, Colectie1)
                        Exit Sub
                    End If



                    If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then

                        Dim String_extra1 As String
                        Dim String_extra2 As String


                        If IsNothing(Rezultat1) = False Then

                            If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                                If IsNothing(Rezultat1) = False Then
                                    For i = 0 To Rezultat1.Value.Count - 1
                                        Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                        Obj1 = Rezultat1.Value.Item(i)
                                        Dim Ent1 As Entity
                                        Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)
                                        If TypeOf Ent1 Is BlockReference Then
                                            Dim Block1 As BlockReference = Ent1

                                            If CheckBox_US_style.Checked = True Then
                                                Dim Continut1 As Double = 99
                                                Dim Continut2 As Double = 0

                                                If Block1.Name.ToUpper.Contains("ELBOW") = False Then


                                                    If Block1.AttributeCollection.Count > 0 Then
                                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                                        Dim Chainage1, Chainage2 As Double
                                                        Dim Length1 As Double

                                                        Dim Chainage_string1 As String
                                                        Dim Chainage_string2 As String
                                                        Dim Is_length As Boolean = False
                                                        Dim Length_string As String
                                                       


                                                        For Each id In attColl
                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                            If attref.Tag.ToUpper = "MAT" Or attref.Tag.ToUpper.Contains("_MAT") = True Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Continut) = True Then
                                                                    Continut1 = CDbl(Continut)
                                                                End If

                                                            End If

                                                            If attref.Tag.ToUpper = "BEGINSTA" Or attref.Tag.ToUpper.Contains("_STABEG") = True Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Chainage1 = Round(CDbl(Replace(Continut, "+", "")), 0)
                                                                    Chainage_string1 = Continut
                                                                End If
                                                            End If

                                                            If attref.Tag.ToUpper = "ENDSTA" Or attref.Tag.ToUpper.Contains("_STAEND") = True Then
                                                                Dim Continut As String = attref.TextString
                                                                If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                    Chainage2 = Round(CDbl(Replace(Continut, "+", "")), 0)
                                                                    Chainage_string2 = Continut
                                                                End If
                                                            End If

                                                            If attref.Tag.ToUpper = "LENGTH" Then
                                                                Dim Continut As String = Replace(attref.TextString, "'", "")
                                                                If IsNumeric(Continut) = True Then
                                                                    Length1 = Round(CDbl(Continut), 0)
                                                                End If
                                                                Length_string = Continut
                                                                Is_length = True
                                                            End If


                                                        Next

                                                        Continut2 = Round(Abs(Chainage1 - Chainage2), 0)

                                                        If Is_length = True Then
                                                            If Not Length1 = Continut2 Then

                                                                If MsgBox("There is a problem with the block " & Block1.Name & vbCrLf & _
                                                                       "Length in block =  " & Length_string & vbCrLf & _
                                                                       "chainage start - end: " & Chainage_string1 & " - " & Chainage_string2 _
                                                                       & vbCrLf & "Do you want it fixed?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                                                    Block1.UpgradeOpen()
                                                                    For Each id In attColl
                                                                        Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)
                                                                        If attref.Tag.ToUpper = "LENGTH" Then
                                                                            attref.TextString = Get_String_Rounded(Abs(Chainage1 - Chainage2), 0) & "'"
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                End If
                                                            End If

                                                        End If

                                                        String_extra1 = Chainage_string1
                                                        String_extra2 = Chainage_string2

                                                    End If


                                                    If Not Continut2 = 0 Then
                                                        ListBox_mat_type.Items.Add(Continut1)
                                                        ListBox_lenghts.Items.Add(Continut2)
                                                        ListBox_chainage1.Items.Add(String_extra1)
                                                        ListBox_chainage2.Items.Add(String_extra2)
                                                        Data_table1.Rows.Add()
                                                        Data_table1.Rows(Index_data_table).Item("TYPE") = Continut1
                                                        Data_table1.Rows(Index_data_table).Item("LENGTH") = Continut2
                                                        Data_table1.Rows(Index_data_table).Item("TRUE_LENGTH") = Continut2

                                                        Index_data_table = Index_data_table + 1
                                                    End If

                                                End If

                                            Else
                                                Dim Continut1 As Double = 99
                                                Dim Continut2 As Double = 0
                                                Dim Length_True As Double = 0
                                                If Block1.Name.ToUpper.Contains("SCREW") = False And Block1.Name.ToUpper.Contains("SAND") = False And Block1.Name.ToUpper.Contains("RIVER") = False _
                                                    And Block1.Name.ToUpper.Contains("CONCRETE") = False And Block1.Name.ToUpper.Contains("CROSSING") = False And Block1.Name.ToUpper.Contains("ANCHOR") = False Then


                                                    If Block1.AttributeCollection.Count > 0 Then
                                                        Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection

                                                        Dim Chainage1, Chainage2 As Double
                                                        Dim Length1 As Double

                                                        Dim Chainage_string1 As String
                                                        Dim Chainage_string2 As String
                                                        Dim Is_length As Boolean = False
                                                        Dim Length_string As String
                                                        Dim ESTE_SCREW As Boolean = False
                                                        For Each id In attColl
                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)
                                                            If attref.Tag.ToUpper = "NO_TYPE" Or attref.Tag.ToUpper = "SPACING" Then
                                                                ESTE_SCREW = True
                                                                Exit For
                                                            End If
                                                        Next

                                                        If ESTE_SCREW = False Then
                                                            For Each id In attColl
                                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)

                                                                If attref.Tag.ToUpper = "MAT" Or attref.Tag.ToUpper.Contains("_MAT") = True Then
                                                                    Dim Continut As String = attref.TextString
                                                                    If IsNumeric(Continut) = True Then
                                                                        Continut1 = CDbl(Continut)
                                                                    End If

                                                                End If

                                                                If attref.Tag.ToUpper = "BEGINSTA" Or attref.Tag.ToUpper.Contains("_STABEG") = True Then
                                                                    Dim Continut As String = attref.TextString
                                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                        Chainage1 = Round(CDbl(Replace(Continut, "+", "")), 2)
                                                                        Chainage_string1 = Continut
                                                                    End If
                                                                End If

                                                                If attref.Tag.ToUpper = "ENDSTA" Or attref.Tag.ToUpper.Contains("_STAEND") = True Then
                                                                    Dim Continut As String = attref.TextString
                                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                                        Chainage2 = Round(CDbl(Replace(Continut, "+", "")), 2)
                                                                        Chainage_string2 = Continut
                                                                    End If
                                                                End If

                                                                If attref.Tag.ToUpper = "LENGTH" Then
                                                                    Dim Continut As String = attref.TextString
                                                                    If IsNumeric(Continut) = True Then
                                                                        Length1 = Round(CDbl(Continut), 2)
                                                                    Else
                                                                        If Continut.Contains("TRUE LENGTH)") = True Then
                                                                            Dim Poz_acol As Integer = InStr(Continut, "(")
                                                                            If Poz_acol > 1 Then
                                                                                Dim string_LEN As String = Mid(Continut, Poz_acol + 1, Len(Continut) - Poz_acol - 13)
                                                                                If IsNumeric(string_LEN) = True Then
                                                                                    Length_True = CDbl(string_LEN)
                                                                                End If
                                                                            End If



                                                                        End If

                                                                    End If
                                                                    Length_string = Continut
                                                                    Is_length = True
                                                                End If


                                                            Next

                                                            Continut2 = Round(Abs(Chainage1 - Chainage2), 2)

                                                            If Is_length = True Then
                                                                If Not Length1 = Continut2 Then

                                                                    If MsgBox("There is a problem with the block " & Block1.Name & vbCrLf & _
                                                                           "Length in block =  " & Length_string & vbCrLf & _
                                                                           "chainage start - end: " & Chainage_string1 & " - " & Chainage_string2 _
                                                                           & vbCrLf & "Do you want it fixed?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                                                        Block1.UpgradeOpen()
                                                                        For Each id In attColl
                                                                            Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForWrite)
                                                                            If attref.Tag.ToUpper = "LENGTH" Then
                                                                                attref.TextString = Get_String_Rounded(Abs(Chainage1 - Chainage2), 1)
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                    End If
                                                                End If

                                                            End If

                                                            String_extra1 = Chainage_string1
                                                            String_extra2 = Chainage_string2
                                                        End If
                                                    End If


                                                    If Not Continut2 = 0 Then
                                                        ListBox_mat_type.Items.Add(Continut1)
                                                        ListBox_lenghts.Items.Add(Continut2)
                                                        ListBox_chainage1.Items.Add(String_extra1)
                                                        ListBox_chainage2.Items.Add(String_extra2)
                                                        Data_table1.Rows.Add()
                                                        Data_table1.Rows(Index_data_table).Item("TYPE") = Continut1
                                                        Data_table1.Rows(Index_data_table).Item("LENGTH") = Continut2
                                                        Data_table1.Rows(Index_data_table).Item("TRUE_LENGTH") = Continut2

                                                        If Not Length_True = 0 Then
                                                            Data_table1.Rows(Index_data_table).Item("TRUE_LENGTH") = Length_True
                                                        End If

                                                        Index_data_table = Index_data_table + 1
                                                    End If

                                                End If
                                            End If

                                        End If


                                    Next
                                    Trans1.Commit()
                                End If
                            End If

                        End If
                    End If

                End Using
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
            Catch ex As Exception
                Editor1.SetImpliedSelection(Empty_array)
                Editor1.WriteMessage(vbLf & "Command:")
                afiseaza_butoanele_pentru_forms(Me, Colectie1)
                MsgBox(ex.Message)
            End Try
            Button_calc_totals_Click(sender, e)
        End Using
    End Sub

    Private Sub Button_calc_dif_Click(sender As Object, e As EventArgs) Handles Button_calc_difference.Click
        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor1 As Autodesk.AutoCAD.EditorInput.Editor
            Editor1 = ThisDrawing.Editor

            Dim Empty_array() As ObjectId
            Editor1.SetImpliedSelection(Empty_array)

            Using lock As DocumentLock = ThisDrawing.LockDocument
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim Rezultat1 As Autodesk.AutoCAD.EditorInput.PromptSelectionResult
                Dim Object_Prompt As New Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                Object_Prompt.MessageForAdding = vbLf & "Select objects:"

                Object_Prompt.SingleOnly = False
                Rezultat1 = Editor1.GetSelection(Object_Prompt)

                If Rezultat1.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
                    If IsNothing(Rezultat1) = False Then
                        Using Trans1 As Autodesk.AutoCAD.DatabaseServices.Transaction = ThisDrawing.TransactionManager.StartTransaction
                            Dim BlockTable1 As Autodesk.AutoCAD.DatabaseServices.BlockTable
                            Dim BTrecord As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                            BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite)

                            TextBox_END_chainage.Text = ""
                            TextBox_BEG_chainage.Text = ""
                            TextBox_diference.Text = ""

                            Dim Chainage1 As Double = 123.123
                            Dim Chainage2 As Double = 123.123
                            Dim Text1 As DBText
                            Dim mText1 As MText
                            Dim Block1 As BlockReference
                            Dim Chainage_start1 As Double = 123.123
                            Dim Chainage_end1 As Double = 123.123
                            Dim Chainage_sta1 As Double = 123.123
                            Dim Block2 As BlockReference
                            Dim Chainage_start2 As Double = 123.123
                            Dim Chainage_end2 As Double = 123.123
                            Dim Chainage_sta2 As Double = 123.123

                            For i = 0 To Rezultat1.Value.Count - 1

                                Dim Obj1 As Autodesk.AutoCAD.EditorInput.SelectedObject
                                Obj1 = Rezultat1.Value.Item(i)
                                Dim Ent1 As Entity
                                Ent1 = Obj1.ObjectId.GetObject(OpenMode.ForRead)

                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                                    Text1 = Ent1
                                    If IsNumeric(Replace(Text1.TextString, "+", "")) = True Then
                                        If Chainage1 = 123.123 Then
                                            Chainage1 = CDbl(Replace(Text1.TextString, "+", ""))
                                        Else
                                            Chainage2 = CDbl(Replace(Text1.TextString, "+", ""))
                                        End If

                                    End If
                                End If
                                If TypeOf Ent1 Is Autodesk.AutoCAD.DatabaseServices.MText Then

                                    mText1 = Ent1
                                    If IsNumeric(Replace(mText1.Text, "+", "")) = True Then
                                        If Chainage1 = 123.123 Then
                                            Chainage1 = CDbl(Replace(mText1.Text, "+", ""))
                                        Else
                                            Chainage2 = CDbl(Replace(mText1.Text, "+", ""))
                                        End If

                                    End If
                                End If


                                If TypeOf Ent1 Is BlockReference Then
                                    If IsNothing(Block1) = True Then
                                        Block1 = Ent1
                                        If Block1.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block1.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)


                                                If attref.Tag.ToUpper = "BEGINSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_start1 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                                If attref.Tag.ToUpper = "ENDSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_end1 = Replace(Continut, "+", "")
                                                    End If
                                                End If
                                                If attref.Tag.ToUpper = "STA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_sta1 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                            Next



                                        End If
                                    Else
                                        Block2 = Ent1
                                        If Block2.AttributeCollection.Count > 0 Then
                                            Dim attColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection = Block2.AttributeCollection
                                            For Each id In attColl
                                                Dim attref As AttributeReference = Trans1.GetObject(id, OpenMode.ForRead)


                                                If attref.Tag.ToUpper = "BEGINSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_start2 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                                If attref.Tag.ToUpper = "ENDSTA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_end2 = Replace(Continut, "+", "")
                                                    End If
                                                End If
                                                If attref.Tag.ToUpper = "STA" Then
                                                    Dim Continut As String = attref.TextString
                                                    If IsNumeric(Replace(Continut, "+", "")) = True Then
                                                        Chainage_sta2 = Replace(Continut, "+", "")
                                                    End If
                                                End If

                                            Next



                                        End If
                                    End If


                                End If

                            Next






                            If Not Chainage1 = 123.123 And Not Chainage2 = 123.123 Then
                                TextBox_diference.Text = Get_String_Rounded(Abs(Chainage1 - Chainage2), 1)
                                TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage1, 1)
                                TextBox_END_chainage.Text = Get_chainage_from_double(Chainage2, 1)
                            End If

                            If (Chainage1 = 123.123 And Not Chainage2 = 123.123) Or (Chainage2 = 123.123 And Not Chainage1 = 123.123) Then
                                Dim PozitieX As Double
                                If IsNothing(Text1) = False Then
                                    PozitieX = Text1.Position.X
                                End If

                                If IsNothing(mText1) = False Then
                                    PozitieX = mText1.Location.X
                                End If

                                Dim PozitieBlockX As Double

                                If IsNothing(Block1) = False Then
                                    PozitieBlockX = Block1.Position.X
                                End If


                                If PozitieX <= PozitieBlockX Then
                                    If Chainage1 = 123.123 Then
                                        If Not Chainage_start1 = 123.123 Then
                                            Chainage1 = Chainage_start1

                                        Else
                                            If Not Chainage_sta1 = 123.123 Then
                                                Chainage1 = Chainage_sta1

                                            Else
                                                If Not Chainage_end1 = 123.123 Then
                                                    Chainage1 = Chainage_end1

                                                Else
                                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                    If Chainage2 = 123.123 Then
                                        If Not Chainage_start1 = 123.123 Then
                                            Chainage2 = Chainage_start1

                                        Else
                                            If Not Chainage_sta1 = 123.123 Then
                                                Chainage2 = Chainage_sta1

                                            Else
                                                If Not Chainage_end1 = 123.123 Then
                                                    Chainage2 = Chainage_end1

                                                Else
                                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If

                                Else
                                    If Chainage1 = 123.123 Then
                                        If Not Chainage_end1 = 123.123 Then
                                            Chainage1 = Chainage_end1
                                        Else
                                            If Not Chainage_sta1 = 123.123 Then
                                                Chainage1 = Chainage_sta1
                                            Else
                                                If Not Chainage_start1 = 123.123 Then
                                                    Chainage1 = Chainage_start1
                                                Else
                                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                    If Chainage2 = 123.123 Then
                                        If Not Chainage_end1 = 123.123 Then
                                            Chainage2 = Chainage_end1
                                        Else
                                            If Not Chainage_sta1 = 123.123 Then
                                                Chainage2 = Chainage_sta1
                                            Else
                                                If Not Chainage_start1 = 123.123 Then
                                                    Chainage2 = Chainage_start1
                                                Else
                                                    afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If



                                TextBox_diference.Text = Get_String_Rounded(Abs(Chainage1 - Chainage2), 1)
                                TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage1, 1)
                                TextBox_END_chainage.Text = Get_chainage_from_double(Chainage2, 1)
                            End If



                            If Chainage1 = 123.123 And Chainage2 = 123.123 Then
                                Dim PozitieX1 As Double
                                Dim PozitieX2 As Double

                                If IsNothing(Block1) = False Then
                                    PozitieX1 = Block1.Position.X
                                End If
                                If IsNothing(Block2) = False Then
                                    PozitieX2 = Block2.Position.X
                                End If

                                If PozitieX1 <= PozitieX2 Then
                                    If Not Chainage_end1 = 123.123 And Not Chainage_start2 = 123.123 Then
                                        Chainage1 = Chainage_end1
                                        Chainage2 = Chainage_start2

                                    Else
                                        If Not Chainage_sta1 = 123.123 Then
                                            Chainage1 = Chainage_sta1
                                            Chainage2 = Chainage_start2

                                        Else
                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    If Not Chainage_end2 = 123.123 And Not Chainage_start1 = 123.123 Then
                                        Chainage1 = Chainage_end2
                                        Chainage2 = Chainage_start1

                                    Else
                                        If Not Chainage_sta2 = 123.123 Then
                                            Chainage1 = Chainage_sta2
                                            Chainage2 = Chainage_start1

                                        Else
                                            If Chainage1 = 123.123 And Chainage2 = 123.123 And Chainage_start2 = 123.123 And Chainage_end2 = 123.123 And Not Chainage_start1 = 123.123 And Not Chainage_end1 = 123.123 Then
                                                Chainage1 = Chainage_start1
                                                Chainage2 = Chainage_end1
                                                TextBox_diference.Text = Get_String_Rounded(Abs(Chainage1 - Chainage2), 1)
                                                TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage1, 1)
                                                TextBox_END_chainage.Text = Get_chainage_from_double(Chainage2, 1)
                                            End If
                                            afiseaza_butoanele_pentru_forms(Me, Colectie1)
                                            Exit Sub
                                        End If
                                    End If

                                End If



                                TextBox_diference.Text = Get_String_Rounded(Abs(Chainage1 - Chainage2), 1)
                                TextBox_BEG_chainage.Text = Get_chainage_from_double(Chainage1, 1)
                                TextBox_END_chainage.Text = Get_chainage_from_double(Chainage2, 1)
                            End If




                            Editor1.Regen()
                            Trans1.Commit()
                        End Using

                    End If
                End If

                ThisDrawing.Editor.WriteMessage(vbLf & "Command:")
            End Using




        Catch ex As Exception
            afiseaza_butoanele_pentru_forms(Me, Colectie1)
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
    End Sub

    Private Sub Button_transfer_to_excel_Click(sender As Object, e As EventArgs) Handles Button_transfer_to_excel.Click
        If IsNumeric(TextBox_ROW_START_XL.Text) = False Then
            Exit Sub
        Else
            Index_row_excel = CInt(TextBox_ROW_START_XL.Text)
        End If


        ascunde_butoanele_pentru_forms(Me, Colectie1)

        Dim W1 As Microsoft.Office.Interop.Excel.Worksheet = Get_active_worksheet_from_Excel()

        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            W1.Range("A" & 1).Value = "Drawing name"
            W1.Range("B" & 1).Value = "Mat 1"
            W1.Range("C" & 1).Value = "Mat 2"
            W1.Range("D" & 1).Value = "Mat 3"
            W1.Range("E" & 1).Value = "Mat 4"
            W1.Range("F" & 1).Value = "Mat 5"
            W1.Range("G" & 1).Value = "Mat 6"
            W1.Range("H" & 1).Value = "Mat 7"
            W1.Range("I" & 1).Value = "Mat 8"
            W1.Range("J" & 1).Value = "Mat 9"
            W1.Range("K" & 1).Value = "Mat 10"
            W1.Range("L" & 1).Value = "Mat 11"
            W1.Range("M" & 1).Value = "Mat 12"
            W1.Range("N" & 1).Value = "ELBOW"

            If Index_row_excel = 1 Then Index_row_excel = 2
            W1.Range("A" & Index_row_excel).Value = ThisDrawing.Name
            W1.Range("B" & Index_row_excel).Value = ""
            W1.Range("C" & Index_row_excel).Value = ""
            W1.Range("D" & Index_row_excel).Value = ""
            W1.Range("E" & Index_row_excel).Value = ""
            W1.Range("F" & Index_row_excel).Value = ""
            W1.Range("G" & Index_row_excel).Value = ""
            W1.Range("H" & Index_row_excel).Value = ""
            W1.Range("I" & Index_row_excel).Value = ""
            W1.Range("J" & Index_row_excel).Value = ""
            W1.Range("K" & Index_row_excel).Value = ""
            W1.Range("L" & Index_row_excel).Value = ""
            W1.Range("M" & Index_row_excel).Value = ""
            W1.Range("N" & Index_row_excel).Value = ""


            If Data_table3.Rows.Count > 0 Then
                For i = 0 To Data_table3.Rows.Count - 1
                    If IsDBNull(Data_table3.Rows(i).Item("TYPE")) = False Then
                        Dim Litera_coloana As String

                        Select Case Round(Data_table3.Rows(i).Item("TYPE"), 0)
                            Case 1
                                Litera_coloana = "B"
                            Case 2
                                Litera_coloana = "C"
                            Case 3
                                Litera_coloana = "D"
                            Case 4
                                Litera_coloana = "E"
                            Case 5
                                Litera_coloana = "F"
                            Case 6
                                Litera_coloana = "G"
                            Case 7
                                Litera_coloana = "H"
                            Case 8
                                Litera_coloana = "I"
                            Case 9
                                Litera_coloana = "J"
                            Case 10
                                Litera_coloana = "K"
                            Case 11
                                Litera_coloana = "L"
                            Case 12
                                Litera_coloana = "M"
                            Case 99
                                Litera_coloana = "N"
                            Case Else
                                Litera_coloana = "AK"
                        End Select
                        W1.Range(Litera_coloana & Index_row_excel).Value = Data_table3.Rows(i).Item("LENGTH")

                    End If
                Next
            End If
            Index_row_excel = Index_row_excel + 1
            TextBox_ROW_START_XL.Text = Index_row_excel.ToString
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele_pentru_forms(Me, Colectie1)
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Command:")

    End Sub

    Private Sub TextBox_diference_click(sender As Object, e As EventArgs) Handles TextBox_diference.Click
        Dim Ch1, Ch2 As Double
        If IsNumeric(Replace(TextBox_BEG_chainage.Text, "+", "")) = True And IsNumeric(Replace(TextBox_END_chainage.Text, "+", "")) = True Then
            Ch1 = Round(CDbl(Replace(TextBox_BEG_chainage.Text, "+", "")), 1)
            Ch2 = Round(CDbl(Replace(TextBox_END_chainage.Text, "+", "")), 1)
            TextBox_diference.Text = Round(Abs(Ch1 - Ch2), 1)

        End If


    End Sub
End Class