
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Transcanada_layers_form
    Dim Data_table_general_layers As System.Data.DataTable
    Dim Data_table_civil_layers As System.Data.DataTable
    Dim Data_table_electrical_layers As System.Data.DataTable
    Dim Data_table_mechanical_layers As System.Data.DataTable
    Dim Data_table_pipeline_layers As System.Data.DataTable
    Dim Data_table_mapping_layers As System.Data.DataTable
    Dim Data_table_extra_layers As System.Data.DataTable
    Dim Data_table_my_list_layers As System.Data.DataTable

    Dim Locatie1 As String = Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) & "\BLOCKS"
    Dim Butoane_visibile As Specialized.StringCollection

    Private Sub ascunde_butoanele()
        Butoane_visibile = New Specialized.StringCollection

        For i = 0 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is Windows.Forms.Button Then
                If Me.Controls(i).Visible = True Then
                    Butoane_visibile.Add(Me.Controls(i).Name)
                    Me.Controls(i).Visible = False
                End If


            End If

        Next

    End Sub
    Private Sub afiseaza_butoanele()
        If Butoane_visibile.Count > 0 Then
            For i = 0 To Me.Controls.Count - 1
                If TypeOf Me.Controls(i) Is Windows.Forms.Button Then
                    If Butoane_visibile.Contains(Me.Controls(i).Name) = True Then
                        Me.Controls(i).Visible = True
                    End If

                End If

            Next
            Butoane_visibile.Clear()

        End If

    End Sub


    Private Sub Transcanada_layers_form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            Data_table_general_layers = New System.Data.DataTable
            Data_table_general_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_general_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_general_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_general_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_general_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_general_layers.Columns.Add("LINETYPE", GetType(String))

            For i = 0 To 20
                Data_table_general_layers.Rows.Add()
            Next


            Data_table_general_layers.Rows(0).Item("NUME_LAYER") = "BALLOONS"
            Data_table_general_layers.Rows(1).Item("NUME_LAYER") = "BOM"
            Data_table_general_layers.Rows(2).Item("NUME_LAYER") = "BORDER"
            Data_table_general_layers.Rows(3).Item("NUME_LAYER") = "BORDERTEXT"
            Data_table_general_layers.Rows(4).Item("NUME_LAYER") = "BUILDINGS"
            Data_table_general_layers.Rows(5).Item("NUME_LAYER") = "CENTERLINES"
            Data_table_general_layers.Rows(6).Item("NUME_LAYER") = "CLOUDS"
            Data_table_general_layers.Rows(7).Item("NUME_LAYER") = "CONSTRUCTION"
            Data_table_general_layers.Rows(8).Item("NUME_LAYER") = "CONTOURS"
            Data_table_general_layers.Rows(9).Item("NUME_LAYER") = "DIMENSION"
            Data_table_general_layers.Rows(10).Item("NUME_LAYER") = "FENCE"
            Data_table_general_layers.Rows(11).Item("NUME_LAYER") = "GIGO"
            Data_table_general_layers.Rows(12).Item("NUME_LAYER") = "GRID"
            Data_table_general_layers.Rows(13).Item("NUME_LAYER") = "HATCH"
            Data_table_general_layers.Rows(14).Item("NUME_LAYER") = "LAND"
            Data_table_general_layers.Rows(15).Item("NUME_LAYER") = "LEGAL"
            Data_table_general_layers.Rows(16).Item("NUME_LAYER") = "LOCATION PLAN"
            Data_table_general_layers.Rows(17).Item("NUME_LAYER") = "MATCH"
            Data_table_general_layers.Rows(18).Item("NUME_LAYER") = "ROAD"
            Data_table_general_layers.Rows(19).Item("NUME_LAYER") = "TEXT"
            Data_table_general_layers.Rows(20).Item("NUME_LAYER") = "UTILITIES"


            Data_table_general_layers.Rows(0).Item("DESCRIERE_LAYER") = "Material ballons"
            Data_table_general_layers.Rows(1).Item("DESCRIERE_LAYER") = "Entire Bill of Material"
            Data_table_general_layers.Rows(2).Item("DESCRIERE_LAYER") = "Drawing Border (linework & Unedited Text)"
            Data_table_general_layers.Rows(3).Item("DESCRIERE_LAYER") = "Drawing Border Atribute Block"
            Data_table_general_layers.Rows(4).Item("DESCRIERE_LAYER") = "Buildings, skid outline, interior walls and partitions, masonry block walls, elevations complete with vents, louvers, windows, doors"
            Data_table_general_layers.Rows(5).Item("DESCRIERE_LAYER") = "Centerlines"
            Data_table_general_layers.Rows(6).Item("DESCRIERE_LAYER") = "Revision clouding, triangles, tie-in points"
            Data_table_general_layers.Rows(7).Item("DESCRIERE_LAYER") = "Construction lines (non-plottable for temporary work only)"
            Data_table_general_layers.Rows(8).Item("DESCRIERE_LAYER") = "Contours and contour elevation text"
            Data_table_general_layers.Rows(9).Item("DESCRIERE_LAYER") = "Dimensions, Coordinates and dimension leader lines"
            Data_table_general_layers.Rows(10).Item("DESCRIERE_LAYER") = "Fence and Gates"
            Data_table_general_layers.Rows(11).Item("DESCRIERE_LAYER") = "Extraneous data from file conversion (Garbage in/Garbage Out)"
            Data_table_general_layers.Rows(12).Item("DESCRIERE_LAYER") = "Grid lines"
            Data_table_general_layers.Rows(13).Item("DESCRIERE_LAYER") = "Hatching and hatching outlines"
            Data_table_general_layers.Rows(14).Item("DESCRIERE_LAYER") = "Land features, trees, berms, ditch,culverts, stock piles, (Civil: trees, berms, ditch, grade, running water, standing water)"
            Data_table_general_layers.Rows(15).Item("DESCRIERE_LAYER") = "Legal boundaries, north arrows, iron pins and all sorts of legal stuff"
            Data_table_general_layers.Rows(16).Item("DESCRIERE_LAYER") = "Location information"
            Data_table_general_layers.Rows(17).Item("DESCRIERE_LAYER") = "Match lines"
            Data_table_general_layers.Rows(18).Item("DESCRIERE_LAYER") = "Road & road allowance, edge of gravel or pavement"
            Data_table_general_layers.Rows(19).Item("DESCRIERE_LAYER") = "All text, notes baloons, section symbols, legends, door-room-wall numbers, text leader lines"
            Data_table_general_layers.Rows(20).Item("DESCRIERE_LAYER") = "Hydro, FOTS, O/H or buried, conduit, telephone, railroads anything that is non Company utilities"


            Data_table_general_layers.Rows(0).Item("CULOARE") = 2
            Data_table_general_layers.Rows(1).Item("CULOARE") = 2
            Data_table_general_layers.Rows(2).Item("CULOARE") = 7
            Data_table_general_layers.Rows(3).Item("CULOARE") = 2
            Data_table_general_layers.Rows(4).Item("CULOARE") = 4
            Data_table_general_layers.Rows(5).Item("CULOARE") = 7
            Data_table_general_layers.Rows(6).Item("CULOARE") = 6
            Data_table_general_layers.Rows(7).Item("CULOARE") = 7
            Data_table_general_layers.Rows(8).Item("CULOARE") = 9
            Data_table_general_layers.Rows(9).Item("CULOARE") = 4
            Data_table_general_layers.Rows(10).Item("CULOARE") = 3
            Data_table_general_layers.Rows(11).Item("CULOARE") = 6
            Data_table_general_layers.Rows(12).Item("CULOARE") = 9
            Data_table_general_layers.Rows(13).Item("CULOARE") = 9
            Data_table_general_layers.Rows(14).Item("CULOARE") = 3
            Data_table_general_layers.Rows(15).Item("CULOARE") = 7
            Data_table_general_layers.Rows(16).Item("CULOARE") = 1
            Data_table_general_layers.Rows(17).Item("CULOARE") = 9
            Data_table_general_layers.Rows(18).Item("CULOARE") = 9
            Data_table_general_layers.Rows(19).Item("CULOARE") = 2
            Data_table_general_layers.Rows(20).Item("CULOARE") = 1

            Data_table_general_layers.Rows(0).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(1).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(2).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_general_layers.Rows(3).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(4).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_general_layers.Rows(5).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(6).Item("LINEWEIGHT") = LineWeight.LineWeight070
            Data_table_general_layers.Rows(7).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(8).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(9).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(10).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(11).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(12).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(13).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(14).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(15).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(16).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(17).Item("LINEWEIGHT") = LineWeight.LineWeight070
            Data_table_general_layers.Rows(18).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(19).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_general_layers.Rows(20).Item("LINEWEIGHT") = LineWeight.LineWeight025

            Data_table_general_layers.Rows(0).Item("PLOT") = True
            Data_table_general_layers.Rows(1).Item("PLOT") = True
            Data_table_general_layers.Rows(2).Item("PLOT") = True
            Data_table_general_layers.Rows(3).Item("PLOT") = True
            Data_table_general_layers.Rows(4).Item("PLOT") = True
            Data_table_general_layers.Rows(5).Item("PLOT") = True
            Data_table_general_layers.Rows(6).Item("PLOT") = True
            Data_table_general_layers.Rows(7).Item("PLOT") = False
            Data_table_general_layers.Rows(8).Item("PLOT") = True
            Data_table_general_layers.Rows(9).Item("PLOT") = True
            Data_table_general_layers.Rows(10).Item("PLOT") = True
            Data_table_general_layers.Rows(11).Item("PLOT") = False
            Data_table_general_layers.Rows(12).Item("PLOT") = True
            Data_table_general_layers.Rows(13).Item("PLOT") = True
            Data_table_general_layers.Rows(14).Item("PLOT") = True
            Data_table_general_layers.Rows(15).Item("PLOT") = True
            Data_table_general_layers.Rows(16).Item("PLOT") = True
            Data_table_general_layers.Rows(17).Item("PLOT") = True
            Data_table_general_layers.Rows(18).Item("PLOT") = True
            Data_table_general_layers.Rows(19).Item("PLOT") = True
            Data_table_general_layers.Rows(20).Item("PLOT") = True


            Data_table_general_layers.Rows(0).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(1).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(2).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(3).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(4).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(5).Item("LINETYPE") = "TCCENTER"
            Data_table_general_layers.Rows(6).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(7).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(8).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(9).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(10).Item("LINETYPE") = "TC_FENCE"
            Data_table_general_layers.Rows(11).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(12).Item("LINETYPE") = "TCDOT4"
            Data_table_general_layers.Rows(13).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(14).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(15).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(16).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(17).Item("LINETYPE") = "TCPHANTOM4"
            Data_table_general_layers.Rows(18).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(19).Item("LINETYPE") = "Continuous"
            Data_table_general_layers.Rows(20).Item("LINETYPE") = "Continuous"




            Data_table_civil_layers = New System.Data.DataTable
            Data_table_civil_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_civil_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_civil_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_civil_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_civil_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_civil_layers.Columns.Add("LINETYPE", GetType(String))

            For i = 0 To 10
                Data_table_civil_layers.Rows.Add()
            Next

            Data_table_civil_layers.Rows(0).Item("NUME_LAYER") = "CARCH"
            Data_table_civil_layers.Rows(1).Item("NUME_LAYER") = "CCONCRETE"
            Data_table_civil_layers.Rows(2).Item("NUME_LAYER") = "CDESIGN"
            Data_table_civil_layers.Rows(3).Item("NUME_LAYER") = "CDRAINAGE"
            Data_table_civil_layers.Rows(4).Item("NUME_LAYER") = "CM STEEL"
            Data_table_civil_layers.Rows(5).Item("NUME_LAYER") = "CPILES"
            Data_table_civil_layers.Rows(6).Item("NUME_LAYER") = "CREBAR"
            Data_table_civil_layers.Rows(7).Item("NUME_LAYER") = "CSSTEEL"
            Data_table_civil_layers.Rows(8).Item("NUME_LAYER") = "CSURVEY"
            Data_table_civil_layers.Rows(9).Item("NUME_LAYER") = "CTRENCH"
            Data_table_civil_layers.Rows(10).Item("NUME_LAYER") = "CYARD"

            Data_table_civil_layers.Rows(0).Item("DESCRIERE_LAYER") = "Counters, sinks, tables & chairs, cabinets, toilets, closets, work benches, insulation, flashing, roof drains, floor drains, fire rating and fire extinguishers, ducts"
            Data_table_civil_layers.Rows(1).Item("DESCRIERE_LAYER") = "Foundation outlines, precast concrete, sections, driveways, miscellaneous concrete, sleepers, anchor blocks"
            Data_table_civil_layers.Rows(2).Item("DESCRIERE_LAYER") = "Design spot elevations or contours"
            Data_table_civil_layers.Rows(3).Item("DESCRIERE_LAYER") = "Catch basins, manholes, sanitary lines, storm lines, industrial lines"
            Data_table_civil_layers.Rows(4).Item("DESCRIERE_LAYER") = "Miscellaneous steel, anchor bolts, grating, door treshhold, welded wire mesh"
            Data_table_civil_layers.Rows(5).Item("DESCRIERE_LAYER") = "Concrete piles, steel piles, pile caps"
            Data_table_civil_layers.Rows(6).Item("DESCRIERE_LAYER") = "Rebar"
            Data_table_civil_layers.Rows(7).Item("DESCRIERE_LAYER") = "Structural steel, pipe racks, Cable trays supports, mechanical trenches, valve/pipe support, girts, purlins"
            Data_table_civil_layers.Rows(8).Item("DESCRIERE_LAYER") = "Boreholes, probeholes, benchmarks, survey monument"
            Data_table_civil_layers.Rows(9).Item("DESCRIERE_LAYER") = "Cable trenches, sidewalks, pre-cast parking curbs, bollards"
            Data_table_civil_layers.Rows(10).Item("DESCRIERE_LAYER") = "Topsoil stockpile, swale, road allowance, culvert"

Data_table_civil_layers.Rows(0).Item("CULOARE") = 5
            Data_table_civil_layers.Rows(1).Item("CULOARE") = 6
            Data_table_civil_layers.Rows(2).Item("CULOARE") = 1
            Data_table_civil_layers.Rows(3).Item("CULOARE") = 4
            Data_table_civil_layers.Rows(4).Item("CULOARE") = 1
            Data_table_civil_layers.Rows(5).Item("CULOARE") = 1
            Data_table_civil_layers.Rows(6).Item("CULOARE") = 1
            Data_table_civil_layers.Rows(7).Item("CULOARE") = 4
            Data_table_civil_layers.Rows(8).Item("CULOARE") = 5
            Data_table_civil_layers.Rows(9).Item("CULOARE") = 1
            Data_table_civil_layers.Rows(10).Item("CULOARE") = 5

            Data_table_civil_layers.Rows(0).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_civil_layers.Rows(1).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_civil_layers.Rows(2).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_civil_layers.Rows(3).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_civil_layers.Rows(4).Item("LINEWEIGHT") = LineWeight.LineWeight070
            Data_table_civil_layers.Rows(5).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_civil_layers.Rows(6).Item("LINEWEIGHT") = LineWeight.LineWeight070
            Data_table_civil_layers.Rows(7).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_civil_layers.Rows(8).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_civil_layers.Rows(9).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_civil_layers.Rows(10).Item("LINEWEIGHT") = LineWeight.LineWeight025

            Data_table_civil_layers.Rows(0).Item("PLOT") = True
            Data_table_civil_layers.Rows(1).Item("PLOT") = True
            Data_table_civil_layers.Rows(2).Item("PLOT") = True
            Data_table_civil_layers.Rows(3).Item("PLOT") = True
            Data_table_civil_layers.Rows(4).Item("PLOT") = True
            Data_table_civil_layers.Rows(5).Item("PLOT") = True
            Data_table_civil_layers.Rows(6).Item("PLOT") = True
            Data_table_civil_layers.Rows(7).Item("PLOT") = True
            Data_table_civil_layers.Rows(8).Item("PLOT") = True
            Data_table_civil_layers.Rows(9).Item("PLOT") = True
            Data_table_civil_layers.Rows(10).Item("PLOT") = True

            Data_table_civil_layers.Rows(0).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(1).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(2).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(3).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(4).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(5).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(6).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(7).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(8).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(9).Item("LINETYPE") = "Continuous"
            Data_table_civil_layers.Rows(10).Item("LINETYPE") = "Continuous"


            Data_table_electrical_layers = New System.Data.DataTable
            Data_table_electrical_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_electrical_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_electrical_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_electrical_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_electrical_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_electrical_layers.Columns.Add("LINETYPE", GetType(String))

            For i = 0 To 11
                Data_table_electrical_layers.Rows.Add()
            Next

            Data_table_electrical_layers.Rows(0).Item("NUME_LAYER") = "EALARM"
            Data_table_electrical_layers.Rows(1).Item("NUME_LAYER") = "EEQUIP"
            Data_table_electrical_layers.Rows(2).Item("NUME_LAYER") = "EGROUND"
            Data_table_electrical_layers.Rows(3).Item("NUME_LAYER") = "EHATCH"
            Data_table_electrical_layers.Rows(4).Item("NUME_LAYER") = "ELIGHT"
            Data_table_electrical_layers.Rows(5).Item("NUME_LAYER") = "EPOWER"
            Data_table_electrical_layers.Rows(6).Item("NUME_LAYER") = "ETRAYS"
            Data_table_electrical_layers.Rows(7).Item("NUME_LAYER") = "EWIRING"
            Data_table_electrical_layers.Rows(8).Item("NUME_LAYER") = "EFWIRE"
            Data_table_electrical_layers.Rows(9).Item("NUME_LAYER") = "ECOM"
            Data_table_electrical_layers.Rows(10).Item("NUME_LAYER") = "ETUBE"
            Data_table_electrical_layers.Rows(11).Item("NUME_LAYER") = "ESYMBOL"

            Data_table_electrical_layers.Rows(0).Item("DESCRIERE_LAYER") = "Fire detection, security, gas, fire, heat, ESD, alarm cables"
            Data_table_electrical_layers.Rows(1).Item("DESCRIERE_LAYER") = "Electrical equipment, I&C, cabinets & panels, telecom cabinet"
            Data_table_electrical_layers.Rows(2).Item("DESCRIERE_LAYER") = "Grounding layout, ground beds, cathodic, ground cables, ground welds, ground symbols"
            Data_table_electrical_layers.Rows(3).Item("DESCRIERE_LAYER") = "Hatch (area classification)"
            Data_table_electrical_layers.Rows(4).Item("DESCRIERE_LAYER") = "Lighting layout, yard lights"
            Data_table_electrical_layers.Rows(5).Item("DESCRIERE_LAYER") = "Building power layout, yard power layout"
            Data_table_electrical_layers.Rows(6).Item("DESCRIERE_LAYER") = "Cable trays"
            Data_table_electrical_layers.Rows(7).Item("DESCRIERE_LAYER") = "Schematics, junction box layouts, loops, single lines, logic diagrams, wiring diagrams"
            Data_table_electrical_layers.Rows(8).Item("DESCRIERE_LAYER") = "Schematics, loops, single lines, logic diagrams, wiring diagrams"
            Data_table_electrical_layers.Rows(9).Item("DESCRIERE_LAYER") = "Telephone wiring, associated symbols, schematics"
            Data_table_electrical_layers.Rows(10).Item("DESCRIERE_LAYER") = "I&C pneumatic tubing and electrical"
            Data_table_electrical_layers.Rows(11).Item("DESCRIERE_LAYER") = "Electrical symbols used in layouts, schematics and diagrams for all"

            Data_table_electrical_layers.Rows(0).Item("CULOARE") = 3
            Data_table_electrical_layers.Rows(1).Item("CULOARE") = 5
            Data_table_electrical_layers.Rows(2).Item("CULOARE") = 4
            Data_table_electrical_layers.Rows(3).Item("CULOARE") = 9
            Data_table_electrical_layers.Rows(4).Item("CULOARE") = 1
            Data_table_electrical_layers.Rows(5).Item("CULOARE") = 6
            Data_table_electrical_layers.Rows(6).Item("CULOARE") = 8
            Data_table_electrical_layers.Rows(7).Item("CULOARE") = 4
            Data_table_electrical_layers.Rows(8).Item("CULOARE") = 2
            Data_table_electrical_layers.Rows(9).Item("CULOARE") = 1
            Data_table_electrical_layers.Rows(10).Item("CULOARE") = 4
            Data_table_electrical_layers.Rows(11).Item("CULOARE") = 6

            Data_table_electrical_layers.Rows(0).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_electrical_layers.Rows(1).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_electrical_layers.Rows(2).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_electrical_layers.Rows(3).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_electrical_layers.Rows(4).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_electrical_layers.Rows(5).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_electrical_layers.Rows(6).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_electrical_layers.Rows(7).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_electrical_layers.Rows(8).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_electrical_layers.Rows(9).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_electrical_layers.Rows(10).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_electrical_layers.Rows(11).Item("LINEWEIGHT") = LineWeight.LineWeight035

            Data_table_electrical_layers.Rows(0).Item("PLOT") = True
            Data_table_electrical_layers.Rows(1).Item("PLOT") = True
            Data_table_electrical_layers.Rows(2).Item("PLOT") = True
            Data_table_electrical_layers.Rows(3).Item("PLOT") = True
            Data_table_electrical_layers.Rows(4).Item("PLOT") = True
            Data_table_electrical_layers.Rows(5).Item("PLOT") = True
            Data_table_electrical_layers.Rows(6).Item("PLOT") = True
            Data_table_electrical_layers.Rows(7).Item("PLOT") = True
            Data_table_electrical_layers.Rows(8).Item("PLOT") = True
            Data_table_electrical_layers.Rows(9).Item("PLOT") = True
            Data_table_electrical_layers.Rows(10).Item("PLOT") = True
            Data_table_electrical_layers.Rows(11).Item("PLOT") = True

            Data_table_electrical_layers.Rows(0).Item("LINETYPE") = "TCHIDDEN2"
            Data_table_electrical_layers.Rows(1).Item("LINETYPE") = "Continuous"
            Data_table_electrical_layers.Rows(2).Item("LINETYPE") = "PHANTOM"
            Data_table_electrical_layers.Rows(3).Item("LINETYPE") = "Continuous"
            Data_table_electrical_layers.Rows(4).Item("LINETYPE") = "TCHIDDEN2"
            Data_table_electrical_layers.Rows(5).Item("LINETYPE") = "TCHIDDEN2"
            Data_table_electrical_layers.Rows(6).Item("LINETYPE") = "Continuous"
            Data_table_electrical_layers.Rows(7).Item("LINETYPE") = "Continuous"
            Data_table_electrical_layers.Rows(8).Item("LINETYPE") = "TCDASH2"
            Data_table_electrical_layers.Rows(9).Item("LINETYPE") = "Continuous"
            Data_table_electrical_layers.Rows(10).Item("LINETYPE") = "Continuous"
            Data_table_electrical_layers.Rows(11).Item("LINETYPE") = "Continuous"


            Data_table_mechanical_layers = New System.Data.DataTable
            Data_table_mechanical_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_mechanical_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_mechanical_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_mechanical_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_mechanical_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_mechanical_layers.Columns.Add("LINETYPE", GetType(String))

            For i = 0 To 30
                Data_table_mechanical_layers.Rows.Add()
            Next
            Data_table_mechanical_layers.Rows(0).Item("NUME_LAYER") = "MEQUIP"
            Data_table_mechanical_layers.Rows(1).Item("NUME_LAYER") = "MHVAC"
            Data_table_mechanical_layers.Rows(2).Item("NUME_LAYER") = "MSAFETY"
            Data_table_mechanical_layers.Rows(3).Item("NUME_LAYER") = "MTOWER"
            Data_table_mechanical_layers.Rows(4).Item("NUME_LAYER") = "M1500LB"
            Data_table_mechanical_layers.Rows(5).Item("NUME_LAYER") = "M150LB"
            Data_table_mechanical_layers.Rows(6).Item("NUME_LAYER") = "M2000LB"
            Data_table_mechanical_layers.Rows(7).Item("NUME_LAYER") = "M2500LB"
            Data_table_mechanical_layers.Rows(8).Item("NUME_LAYER") = "M3000LB"
            Data_table_mechanical_layers.Rows(9).Item("NUME_LAYER") = "M300LB"
            Data_table_mechanical_layers.Rows(10).Item("NUME_LAYER") = "M4000LB"
            Data_table_mechanical_layers.Rows(11).Item("NUME_LAYER") = "M400LB"
            Data_table_mechanical_layers.Rows(12).Item("NUME_LAYER") = "M6000LB"
            Data_table_mechanical_layers.Rows(13).Item("NUME_LAYER") = "M600LB"
            Data_table_mechanical_layers.Rows(14).Item("NUME_LAYER") = "M900LB"
            Data_table_mechanical_layers.Rows(15).Item("NUME_LAYER") = "M5000LB"
            Data_table_mechanical_layers.Rows(16).Item("NUME_LAYER") = "M10000LB"
            Data_table_mechanical_layers.Rows(17).Item("NUME_LAYER") = "MEXIST"
            Data_table_mechanical_layers.Rows(18).Item("NUME_LAYER") = "MATTRIB"
            Data_table_mechanical_layers.Rows(19).Item("NUME_LAYER") = "MCONTIN"
            Data_table_mechanical_layers.Rows(20).Item("NUME_LAYER") = "MDASH"
            Data_table_mechanical_layers.Rows(21).Item("NUME_LAYER") = "MDOT"
            Data_table_mechanical_layers.Rows(22).Item("NUME_LAYER") = "MHIDDEN"
            Data_table_mechanical_layers.Rows(23).Item("NUME_LAYER") = "MINSTR"
            Data_table_mechanical_layers.Rows(24).Item("NUME_LAYER") = "MMISC"
            Data_table_mechanical_layers.Rows(25).Item("NUME_LAYER") = "MPHANTOM"
            Data_table_mechanical_layers.Rows(26).Item("NUME_LAYER") = "MPROCESS"
            Data_table_mechanical_layers.Rows(27).Item("NUME_LAYER") = "MTUBING"
            Data_table_mechanical_layers.Rows(28).Item("NUME_LAYER") = "MUTILITY"
            Data_table_mechanical_layers.Rows(29).Item("NUME_LAYER") = "MVALVES"
            Data_table_mechanical_layers.Rows(30).Item("NUME_LAYER") = "MFLANGES"

            Data_table_mechanical_layers.Rows(0).Item("DESCRIERE_LAYER") = "Heaters, vessels, scrubbers, jets, air conditioners, etc."
            Data_table_mechanical_layers.Rows(1).Item("DESCRIERE_LAYER") = "Return and supply"
            Data_table_mechanical_layers.Rows(2).Item("DESCRIERE_LAYER") = "Fire and safety symbols, wind sock"
            Data_table_mechanical_layers.Rows(3).Item("DESCRIERE_LAYER") = "Communications, sattelite dishes"
            Data_table_mechanical_layers.Rows(4).Item("DESCRIERE_LAYER") = "Piping with 1500 lb rating"
            Data_table_mechanical_layers.Rows(5).Item("DESCRIERE_LAYER") = "Piping with 150 lb rating"
            Data_table_mechanical_layers.Rows(6).Item("DESCRIERE_LAYER") = "Piping with 2000 lb rating"
            Data_table_mechanical_layers.Rows(7).Item("DESCRIERE_LAYER") = "Piping with 2500 lb rating"
            Data_table_mechanical_layers.Rows(8).Item("DESCRIERE_LAYER") = "Piping with 3000 lb rating"
            Data_table_mechanical_layers.Rows(9).Item("DESCRIERE_LAYER") = "Piping with 300 lb rating"
            Data_table_mechanical_layers.Rows(10).Item("DESCRIERE_LAYER") = "Piping with 4000 lb rating"
            Data_table_mechanical_layers.Rows(11).Item("DESCRIERE_LAYER") = "Piping with 400 lb rating"
            Data_table_mechanical_layers.Rows(12).Item("DESCRIERE_LAYER") = "Piping with 6000 lb rating"
            Data_table_mechanical_layers.Rows(13).Item("DESCRIERE_LAYER") = "Piping with 600 lb rating"
            Data_table_mechanical_layers.Rows(14).Item("DESCRIERE_LAYER") = "Piping with 900 lb rating"
            Data_table_mechanical_layers.Rows(15).Item("DESCRIERE_LAYER") = "Piping with 5000 lb rating"
            Data_table_mechanical_layers.Rows(16).Item("DESCRIERE_LAYER") = "Piping with 10000 lb rating"
            Data_table_mechanical_layers.Rows(17).Item("DESCRIERE_LAYER") = "Piping existing"
            Data_table_mechanical_layers.Rows(18).Item("DESCRIERE_LAYER") = "Piping attributes"
            Data_table_mechanical_layers.Rows(19).Item("DESCRIERE_LAYER") = "Piping continuous lines"
            Data_table_mechanical_layers.Rows(20).Item("DESCRIERE_LAYER") = "Piping dashed lines"
            Data_table_mechanical_layers.Rows(21).Item("DESCRIERE_LAYER") = "Piping dotted lines"
            Data_table_mechanical_layers.Rows(22).Item("DESCRIERE_LAYER") = "Piping hidden lines"
            Data_table_mechanical_layers.Rows(23).Item("DESCRIERE_LAYER") = "Piping instrumentation symbols"
            Data_table_mechanical_layers.Rows(24).Item("DESCRIERE_LAYER") = "Piping miscellaneous"
            Data_table_mechanical_layers.Rows(25).Item("DESCRIERE_LAYER") = "Piping phantom lines"
            Data_table_mechanical_layers.Rows(26).Item("DESCRIERE_LAYER") = "Piping and Instrumentation Diagram Piping Main Process Lines"
            Data_table_mechanical_layers.Rows(27).Item("DESCRIERE_LAYER") = "Piping Tubing"
            Data_table_mechanical_layers.Rows(28).Item("DESCRIERE_LAYER") = "Piping and Instrumentation Diagram Piping Secondary Process Lines"
            Data_table_mechanical_layers.Rows(29).Item("DESCRIERE_LAYER") = "Piping and Instrumentation Diagram Piping Valves"
            Data_table_mechanical_layers.Rows(30).Item("DESCRIERE_LAYER") = "Piping and Instrumentation Diagram Piping Flanges"

            Data_table_mechanical_layers.Rows(0).Item("CULOARE") = 6
            Data_table_mechanical_layers.Rows(1).Item("CULOARE") = 4
            Data_table_mechanical_layers.Rows(2).Item("CULOARE") = 1
            Data_table_mechanical_layers.Rows(3).Item("CULOARE") = 3
            Data_table_mechanical_layers.Rows(4).Item("CULOARE") = 1
            Data_table_mechanical_layers.Rows(5).Item("CULOARE") = 2
            Data_table_mechanical_layers.Rows(6).Item("CULOARE") = 1
            Data_table_mechanical_layers.Rows(7).Item("CULOARE") = 2
            Data_table_mechanical_layers.Rows(8).Item("CULOARE") = 2
            Data_table_mechanical_layers.Rows(9).Item("CULOARE") = 3
            Data_table_mechanical_layers.Rows(10).Item("CULOARE") = 4
            Data_table_mechanical_layers.Rows(11).Item("CULOARE") = 4
            Data_table_mechanical_layers.Rows(12).Item("CULOARE") = 3
            Data_table_mechanical_layers.Rows(13).Item("CULOARE") = 1
            Data_table_mechanical_layers.Rows(14).Item("CULOARE") = 6
            Data_table_mechanical_layers.Rows(15).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(16).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(17).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(18).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(19).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(20).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(21).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(22).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(23).Item("CULOARE") = 2
            Data_table_mechanical_layers.Rows(24).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(25).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(26).Item("CULOARE") = 4
            Data_table_mechanical_layers.Rows(27).Item("CULOARE") = 7
            Data_table_mechanical_layers.Rows(28).Item("CULOARE") = 2
            Data_table_mechanical_layers.Rows(29).Item("CULOARE") = 1
            Data_table_mechanical_layers.Rows(30).Item("CULOARE") = 7

            Data_table_mechanical_layers.Rows(0).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(1).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(2).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(3).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(4).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(5).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(6).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(7).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(8).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(9).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(10).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(11).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(12).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(13).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(14).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(15).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(16).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(17).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(18).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(19).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(20).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(21).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(22).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(23).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(24).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(25).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(26).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mechanical_layers.Rows(27).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mechanical_layers.Rows(28).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(29).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mechanical_layers.Rows(30).Item("LINEWEIGHT") = LineWeight.LineWeight025

            Data_table_mechanical_layers.Rows(0).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(1).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(2).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(3).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(4).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(5).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(6).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(7).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(8).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(9).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(10).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(11).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(12).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(13).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(14).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(15).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(16).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(17).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(18).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(19).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(20).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(21).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(22).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(23).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(24).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(25).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(26).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(27).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(28).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(29).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(30).Item("PLOT") = True
            Data_table_mechanical_layers.Rows(0).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(1).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(2).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(3).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(4).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(5).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(6).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(7).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(8).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(9).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(10).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(11).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(12).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(13).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(14).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(15).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(16).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(17).Item("LINETYPE") = "TCHIDDEN2"
            Data_table_mechanical_layers.Rows(18).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(19).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(20).Item("LINETYPE") = "TCHIDDEN2"
            Data_table_mechanical_layers.Rows(21).Item("LINETYPE") = "TCDOT2"
            Data_table_mechanical_layers.Rows(22).Item("LINETYPE") = "TCDASH2"
            Data_table_mechanical_layers.Rows(23).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(24).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(25).Item("LINETYPE") = "TCPHANTOM4"
            Data_table_mechanical_layers.Rows(26).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(27).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(28).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(29).Item("LINETYPE") = "Continuous"
            Data_table_mechanical_layers.Rows(30).Item("LINETYPE") = "Continuous"



            Data_table_pipeline_layers = New System.Data.DataTable
            Data_table_pipeline_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_pipeline_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_pipeline_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_pipeline_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_pipeline_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_pipeline_layers.Columns.Add("LINETYPE", GetType(String))

            For i = 0 To 10
                Data_table_pipeline_layers.Rows.Add()
            Next
            Data_table_pipeline_layers.Rows(0).Item("NUME_LAYER") = "PENVIRO"
            Data_table_pipeline_layers.Rows(1).Item("NUME_LAYER") = "PGRADE"
            Data_table_pipeline_layers.Rows(2).Item("NUME_LAYER") = "PWATER"
            Data_table_pipeline_layers.Rows(3).Item("NUME_LAYER") = "PGRID"
            Data_table_pipeline_layers.Rows(4).Item("NUME_LAYER") = "PRW"
            Data_table_pipeline_layers.Rows(5).Item("NUME_LAYER") = "PSOIL"
            Data_table_pipeline_layers.Rows(6).Item("NUME_LAYER") = "PTEXT"
            Data_table_pipeline_layers.Rows(7).Item("NUME_LAYER") = "PWORK"
            Data_table_pipeline_layers.Rows(8).Item("NUME_LAYER") = "PEXIST"
            Data_table_pipeline_layers.Rows(9).Item("NUME_LAYER") = "PNEW"
            Data_table_pipeline_layers.Rows(10).Item("NUME_LAYER") = "PFNEW"

            Data_table_pipeline_layers.Rows(0).Item("DESCRIERE_LAYER") = "Results of soil testing, including text and linework"
            Data_table_pipeline_layers.Rows(1).Item("DESCRIERE_LAYER") = "Grade"
            Data_table_pipeline_layers.Rows(2).Item("DESCRIERE_LAYER") = "Water"
            Data_table_pipeline_layers.Rows(3).Item("DESCRIERE_LAYER") = "Profile grids"
            Data_table_pipeline_layers.Rows(4).Item("DESCRIERE_LAYER") = "Right of way limits"
            Data_table_pipeline_layers.Rows(5).Item("DESCRIERE_LAYER") = "Soil testing information and soil evaluation (symbols & text)"
            Data_table_pipeline_layers.Rows(6).Item("DESCRIERE_LAYER") = "EGIS input, alignment sheets"
            Data_table_pipeline_layers.Rows(7).Item("DESCRIERE_LAYER") = "Safe working boundary limits"
            Data_table_pipeline_layers.Rows(8).Item("DESCRIERE_LAYER") = "Existing piping for Pipeline Facilities"
            Data_table_pipeline_layers.Rows(9).Item("DESCRIERE_LAYER") = "New piping for pipeline"
            Data_table_pipeline_layers.Rows(10).Item("DESCRIERE_LAYER") = "New piping for pipeline facilities"

            Data_table_pipeline_layers.Rows(0).Item("CULOARE") = 7
            Data_table_pipeline_layers.Rows(1).Item("CULOARE") = 3
            Data_table_pipeline_layers.Rows(2).Item("CULOARE") = 5
            Data_table_pipeline_layers.Rows(3).Item("CULOARE") = 6
            Data_table_pipeline_layers.Rows(4).Item("CULOARE") = 3
            Data_table_pipeline_layers.Rows(5).Item("CULOARE") = 7
            Data_table_pipeline_layers.Rows(6).Item("CULOARE") = 3
            Data_table_pipeline_layers.Rows(7).Item("CULOARE") = 3
            Data_table_pipeline_layers.Rows(8).Item("CULOARE") = 7
            Data_table_pipeline_layers.Rows(9).Item("CULOARE") = 7
            Data_table_pipeline_layers.Rows(10).Item("CULOARE") = 7

            Data_table_pipeline_layers.Rows(0).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_pipeline_layers.Rows(1).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_pipeline_layers.Rows(2).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_pipeline_layers.Rows(3).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_pipeline_layers.Rows(4).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_pipeline_layers.Rows(5).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_pipeline_layers.Rows(6).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_pipeline_layers.Rows(7).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_pipeline_layers.Rows(8).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_pipeline_layers.Rows(9).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_pipeline_layers.Rows(10).Item("LINEWEIGHT") = LineWeight.LineWeight035

            Data_table_pipeline_layers.Rows(0).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(1).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(2).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(3).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(4).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(5).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(6).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(7).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(8).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(9).Item("PLOT") = True
            Data_table_pipeline_layers.Rows(10).Item("PLOT") = True

            Data_table_pipeline_layers.Rows(0).Item("LINETYPE") = "Continuous"
            Data_table_pipeline_layers.Rows(1).Item("LINETYPE") = "Continuous"
            Data_table_pipeline_layers.Rows(2).Item("LINETYPE") = "Continuous"
            Data_table_pipeline_layers.Rows(3).Item("LINETYPE") = "TCHIDDEN"
            Data_table_pipeline_layers.Rows(4).Item("LINETYPE") = "Continuous"
            Data_table_pipeline_layers.Rows(5).Item("LINETYPE") = "Continuous"
            Data_table_pipeline_layers.Rows(6).Item("LINETYPE") = "Continuous"
            Data_table_pipeline_layers.Rows(7).Item("LINETYPE") = "TCPHANTOM4"
            Data_table_pipeline_layers.Rows(8).Item("LINETYPE") = "TCHIDDEN2"
            Data_table_pipeline_layers.Rows(9).Item("LINETYPE") = "Continuous"
            Data_table_pipeline_layers.Rows(10).Item("LINETYPE") = "Continuous"


            Data_table_mapping_layers = New System.Data.DataTable
            Data_table_mapping_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_mapping_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_mapping_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_mapping_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_mapping_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_mapping_layers.Columns.Add("LINETYPE", GetType(String))

            For i = 0 To 174
                Data_table_mapping_layers.Rows.Add()
            Next

            Data_table_mapping_layers.Rows(0).Item("NUME_LAYER") = "100-1"
            Data_table_mapping_layers.Rows(1).Item("NUME_LAYER") = "100-2"
            Data_table_mapping_layers.Rows(2).Item("NUME_LAYER") = "100-3"
            Data_table_mapping_layers.Rows(3).Item("NUME_LAYER") = "100-4"
            Data_table_mapping_layers.Rows(4).Item("NUME_LAYER") = "100-5"
            Data_table_mapping_layers.Rows(5).Item("NUME_LAYER") = "100-6"
            Data_table_mapping_layers.Rows(6).Item("NUME_LAYER") = "100-7"
            Data_table_mapping_layers.Rows(7).Item("NUME_LAYER") = "1100-1"
            Data_table_mapping_layers.Rows(8).Item("NUME_LAYER") = "1200-1"
            Data_table_mapping_layers.Rows(9).Item("NUME_LAYER") = "1200-2"
            Data_table_mapping_layers.Rows(10).Item("NUME_LAYER") = "1300-1"
            Data_table_mapping_layers.Rows(11).Item("NUME_LAYER") = "1300-2"
            Data_table_mapping_layers.Rows(12).Item("NUME_LAYER") = "1400-1"
            Data_table_mapping_layers.Rows(13).Item("NUME_LAYER") = "1400-2"
            Data_table_mapping_layers.Rows(14).Item("NUME_LAYER") = "1600-1"
            Data_table_mapping_layers.Rows(15).Item("NUME_LAYER") = "1700-1"
            Data_table_mapping_layers.Rows(16).Item("NUME_LAYER") = "200-1"
            Data_table_mapping_layers.Rows(17).Item("NUME_LAYER") = "200-2"
            Data_table_mapping_layers.Rows(18).Item("NUME_LAYER") = "200-3"
            Data_table_mapping_layers.Rows(19).Item("NUME_LAYER") = "300-1"
            Data_table_mapping_layers.Rows(20).Item("NUME_LAYER") = "300-2"
            Data_table_mapping_layers.Rows(21).Item("NUME_LAYER") = "400-1"
            Data_table_mapping_layers.Rows(22).Item("NUME_LAYER") = "400-2"
            Data_table_mapping_layers.Rows(23).Item("NUME_LAYER") = "400-3"
            Data_table_mapping_layers.Rows(24).Item("NUME_LAYER") = "500-1"
            Data_table_mapping_layers.Rows(25).Item("NUME_LAYER") = "500-2"
            Data_table_mapping_layers.Rows(26).Item("NUME_LAYER") = "500-3"
            Data_table_mapping_layers.Rows(27).Item("NUME_LAYER") = "700-1"
            Data_table_mapping_layers.Rows(28).Item("NUME_LAYER") = "700-2"
            Data_table_mapping_layers.Rows(29).Item("NUME_LAYER") = "800-1"
            Data_table_mapping_layers.Rows(30).Item("NUME_LAYER") = "800-2"
            Data_table_mapping_layers.Rows(31).Item("NUME_LAYER") = "900-1"
            Data_table_mapping_layers.Rows(32).Item("NUME_LAYER") = "900-2"
            Data_table_mapping_layers.Rows(33).Item("NUME_LAYER") = "ALBERTA_COMP_STN"
            Data_table_mapping_layers.Rows(34).Item("NUME_LAYER") = "ALBERTA_DELIVERY_SALES_STN"
            Data_table_mapping_layers.Rows(35).Item("NUME_LAYER") = "ALBERTA_FUTURE_PIPELINE"
            Data_table_mapping_layers.Rows(36).Item("NUME_LAYER") = "ALBERTA_PIPELINE"
            Data_table_mapping_layers.Rows(37).Item("NUME_LAYER") = "ALBERTA_PROPOSE"
            Data_table_mapping_layers.Rows(38).Item("NUME_LAYER") = "ALBERTA_RECEIPT_METER_STN"
            Data_table_mapping_layers.Rows(39).Item("NUME_LAYER") = "BDRY_AIRSTRIP"
            Data_table_mapping_layers.Rows(40).Item("NUME_LAYER") = "BDRY_CADASTRAL"
            Data_table_mapping_layers.Rows(41).Item("NUME_LAYER") = "BDRY_DISTRICT"
            Data_table_mapping_layers.Rows(42).Item("NUME_LAYER") = "BDRY_FED"
            Data_table_mapping_layers.Rows(43).Item("NUME_LAYER") = "BDRY_FOREST_RESERVE"
            Data_table_mapping_layers.Rows(44).Item("NUME_LAYER") = "BDRY_FOTS"
            Data_table_mapping_layers.Rows(45).Item("NUME_LAYER") = "BDRY_INDIAN_RESERVE"
            Data_table_mapping_layers.Rows(46).Item("NUME_LAYER") = "BDRY_INTERNATIONAL"
            Data_table_mapping_layers.Rows(47).Item("NUME_LAYER") = "BDRY_LOT_SUR"
            Data_table_mapping_layers.Rows(48).Item("NUME_LAYER") = "BDRY_LOT_UNS"
            Data_table_mapping_layers.Rows(49).Item("NUME_LAYER") = "BDRY_METIS_SETTLEMENT"
            Data_table_mapping_layers.Rows(50).Item("NUME_LAYER") = "BDRY_MILITARY_RESERVE"
            Data_table_mapping_layers.Rows(51).Item("NUME_LAYER") = "BDRY_OPERATIONS_REGIONS"
            Data_table_mapping_layers.Rows(52).Item("NUME_LAYER") = "BDRY_PARK"
            Data_table_mapping_layers.Rows(53).Item("NUME_LAYER") = "BDRY_PROPOSED_SUBDIVISION"
            Data_table_mapping_layers.Rows(54).Item("NUME_LAYER") = "BDRY_PROVINCE"
            Data_table_mapping_layers.Rows(55).Item("NUME_LAYER") = "BDRY_ROAD_ALLOW"
            Data_table_mapping_layers.Rows(56).Item("NUME_LAYER") = "BDRY_ROWL"
            Data_table_mapping_layers.Rows(57).Item("NUME_LAYER") = "BDRY_STATE"
            Data_table_mapping_layers.Rows(58).Item("NUME_LAYER") = "BDRY_TCPL"
            Data_table_mapping_layers.Rows(59).Item("NUME_LAYER") = "BDRY_TCPL_PROP_ROWL"
            Data_table_mapping_layers.Rows(60).Item("NUME_LAYER") = "BDRY_TCPL_ROWL"
            Data_table_mapping_layers.Rows(61).Item("NUME_LAYER") = "BDRY_TCPL_SMONU"
            Data_table_mapping_layers.Rows(62).Item("NUME_LAYER") = "BDRY_TCPL_WKROOM"
            Data_table_mapping_layers.Rows(63).Item("NUME_LAYER") = "BDRY_TWP"
            Data_table_mapping_layers.Rows(64).Item("NUME_LAYER") = "BORDER"
            Data_table_mapping_layers.Rows(65).Item("NUME_LAYER") = "BUILDING"
            Data_table_mapping_layers.Rows(66).Item("NUME_LAYER") = "BUILDING_OFF"
            Data_table_mapping_layers.Rows(67).Item("NUME_LAYER") = "CENTERLINE"
            Data_table_mapping_layers.Rows(68).Item("NUME_LAYER") = "CITY"
            Data_table_mapping_layers.Rows(69).Item("NUME_LAYER") = "CONTOUR"
            Data_table_mapping_layers.Rows(70).Item("NUME_LAYER") = "CONTROL"
            Data_table_mapping_layers.Rows(71).Item("NUME_LAYER") = "CUT_LINE"
            Data_table_mapping_layers.Rows(72).Item("NUME_LAYER") = "DAM"
            Data_table_mapping_layers.Rows(73).Item("NUME_LAYER") = "DAM_BEAVER"
            Data_table_mapping_layers.Rows(74).Item("NUME_LAYER") = "DIM"
            Data_table_mapping_layers.Rows(75).Item("NUME_LAYER") = "DITCH"
            Data_table_mapping_layers.Rows(76).Item("NUME_LAYER") = "ELEVATION"
            Data_table_mapping_layers.Rows(77).Item("NUME_LAYER") = "FAC_FENCE"
            Data_table_mapping_layers.Rows(78).Item("NUME_LAYER") = "FLOW_ARROW"
            Data_table_mapping_layers.Rows(79).Item("NUME_LAYER") = "FOOTHILLS_KP"
            Data_table_mapping_layers.Rows(80).Item("NUME_LAYER") = "FOOTHILLS_PIPELINE"
            Data_table_mapping_layers.Rows(81).Item("NUME_LAYER") = "FOOTHILLS_ROWL"
            Data_table_mapping_layers.Rows(82).Item("NUME_LAYER") = "FOOTHILLS_SMS"
            Data_table_mapping_layers.Rows(83).Item("NUME_LAYER") = "FOOTHILLS_TEXT"
            Data_table_mapping_layers.Rows(84).Item("NUME_LAYER") = "FOOTHILLS_VALVE"
            Data_table_mapping_layers.Rows(85).Item("NUME_LAYER") = "FOREIGN_PIPE"
            Data_table_mapping_layers.Rows(86).Item("NUME_LAYER") = "FOTS"
            Data_table_mapping_layers.Rows(87).Item("NUME_LAYER") = "GLACIAL_LIMIT"
            Data_table_mapping_layers.Rows(88).Item("NUME_LAYER") = "GRADE"
            Data_table_mapping_layers.Rows(89).Item("NUME_LAYER") = "GRID_ALBERTA"
            Data_table_mapping_layers.Rows(90).Item("NUME_LAYER") = "GRID_BC"
            Data_table_mapping_layers.Rows(91).Item("NUME_LAYER") = "GRID_NTS"
            Data_table_mapping_layers.Rows(92).Item("NUME_LAYER") = "GRID_SASKATCHEWAN"
            Data_table_mapping_layers.Rows(93).Item("NUME_LAYER") = "GRID-MAP"
            Data_table_mapping_layers.Rows(94).Item("NUME_LAYER") = "HWY_CENTERLINE"
            Data_table_mapping_layers.Rows(95).Item("NUME_LAYER") = "ICE_LIMIT"
            Data_table_mapping_layers.Rows(96).Item("NUME_LAYER") = "IMAGE"
            Data_table_mapping_layers.Rows(97).Item("NUME_LAYER") = "KP"
            Data_table_mapping_layers.Rows(98).Item("NUME_LAYER") = "LEGEND"
            Data_table_mapping_layers.Rows(99).Item("NUME_LAYER") = "LINE_1"
            Data_table_mapping_layers.Rows(100).Item("NUME_LAYER") = "LINE_2"
            Data_table_mapping_layers.Rows(101).Item("NUME_LAYER") = "LINE_3"
            Data_table_mapping_layers.Rows(102).Item("NUME_LAYER") = "LINE_4"
            Data_table_mapping_layers.Rows(103).Item("NUME_LAYER") = "LINE_5"
            Data_table_mapping_layers.Rows(104).Item("NUME_LAYER") = "LINE_6"
            Data_table_mapping_layers.Rows(105).Item("NUME_LAYER") = "LINE_7"
            Data_table_mapping_layers.Rows(106).Item("NUME_LAYER") = "LINE_MEDIUM"
            Data_table_mapping_layers.Rows(107).Item("NUME_LAYER") = "LINE_THICK"
            Data_table_mapping_layers.Rows(108).Item("NUME_LAYER") = "LINE_THIN"
            Data_table_mapping_layers.Rows(109).Item("NUME_LAYER") = "LOGO"
            Data_table_mapping_layers.Rows(110).Item("NUME_LAYER") = "MLV"
            Data_table_mapping_layers.Rows(111).Item("NUME_LAYER") = "NEATLINE"
            Data_table_mapping_layers.Rows(112).Item("NUME_LAYER") = "NORTH_ARROW"
            Data_table_mapping_layers.Rows(113).Item("NUME_LAYER") = "OFF"
            Data_table_mapping_layers.Rows(114).Item("NUME_LAYER") = "PIPELINE_FOREIGN"
            Data_table_mapping_layers.Rows(115).Item("NUME_LAYER") = "PIPELINE_GAS"
            Data_table_mapping_layers.Rows(116).Item("NUME_LAYER") = "PIPELINE_MULTIUSE"
            Data_table_mapping_layers.Rows(117).Item("NUME_LAYER") = "PIPELINE_OIL"
            Data_table_mapping_layers.Rows(118).Item("NUME_LAYER") = "RAILWAY"
            Data_table_mapping_layers.Rows(119).Item("NUME_LAYER") = "RAILWAY_ABANDONED"
            Data_table_mapping_layers.Rows(120).Item("NUME_LAYER") = "RAILWAY_MULTIPLE"
            Data_table_mapping_layers.Rows(121).Item("NUME_LAYER") = "REGISTRATION_MARKS"
            Data_table_mapping_layers.Rows(122).Item("NUME_LAYER") = "ROAD_CENTRELINE"
            Data_table_mapping_layers.Rows(123).Item("NUME_LAYER") = "ROAD_DIVIDED"
            Data_table_mapping_layers.Rows(124).Item("NUME_LAYER") = "ROAD_GRAVEL"
            Data_table_mapping_layers.Rows(125).Item("NUME_LAYER") = "ROAD_HIGHWAY"
            Data_table_mapping_layers.Rows(126).Item("NUME_LAYER") = "ROAD_PAVED"
            Data_table_mapping_layers.Rows(127).Item("NUME_LAYER") = "ROAD_SECONDARY"
            Data_table_mapping_layers.Rows(128).Item("NUME_LAYER") = "ROAD_TRAIL"
            Data_table_mapping_layers.Rows(129).Item("NUME_LAYER") = "ROAD_WINTER"
            Data_table_mapping_layers.Rows(130).Item("NUME_LAYER") = "SEISMIC_LINE"
            Data_table_mapping_layers.Rows(131).Item("NUME_LAYER") = "SCALEBAR"
            Data_table_mapping_layers.Rows(132).Item("NUME_LAYER") = "SPOT_HEIGHT"
            Data_table_mapping_layers.Rows(133).Item("NUME_LAYER") = "SYMBOL"
            Data_table_mapping_layers.Rows(134).Item("NUME_LAYER") = "SYMBOL_OFF"
            Data_table_mapping_layers.Rows(135).Item("NUME_LAYER") = "SYMBOL_ROAD"
            Data_table_mapping_layers.Rows(136).Item("NUME_LAYER") = "TCPL_CATHODIC"
            Data_table_mapping_layers.Rows(137).Item("NUME_LAYER") = "TCPL_COMP_STN"
            Data_table_mapping_layers.Rows(138).Item("NUME_LAYER") = "TCPL_FABRICATION"
            Data_table_mapping_layers.Rows(139).Item("NUME_LAYER") = "TCPL_PIPELINE"
            Data_table_mapping_layers.Rows(140).Item("NUME_LAYER") = "TCPL_PROPOSED_PIPELINE"
            Data_table_mapping_layers.Rows(141).Item("NUME_LAYER") = "TCPL_REPLACEMENT"
            Data_table_mapping_layers.Rows(142).Item("NUME_LAYER") = "TCPL_SMS"
            Data_table_mapping_layers.Rows(143).Item("NUME_LAYER") = "TCPL_TEXT"
            Data_table_mapping_layers.Rows(144).Item("NUME_LAYER") = "TCPL_VALVE"
            Data_table_mapping_layers.Rows(145).Item("NUME_LAYER") = "TEXT_BUILDING"
            Data_table_mapping_layers.Rows(146).Item("NUME_LAYER") = "TEXT_CADASTRAL"
            Data_table_mapping_layers.Rows(147).Item("NUME_LAYER") = "TEXT_CITY"
            Data_table_mapping_layers.Rows(148).Item("NUME_LAYER") = "TEXT_CONTOUR"
            Data_table_mapping_layers.Rows(149).Item("NUME_LAYER") = "TEXT_FACILITY"
            Data_table_mapping_layers.Rows(150).Item("NUME_LAYER") = "TEXT_FED"
            Data_table_mapping_layers.Rows(151).Item("NUME_LAYER") = "TEXT_GENERAL"
            Data_table_mapping_layers.Rows(152).Item("NUME_LAYER") = "TEXT_GRID"
            Data_table_mapping_layers.Rows(153).Item("NUME_LAYER") = "TEXT_LEGEND"
            Data_table_mapping_layers.Rows(154).Item("NUME_LAYER") = "TEXT_PARK"
            Data_table_mapping_layers.Rows(155).Item("NUME_LAYER") = "TEXT_PIPELINE"
            Data_table_mapping_layers.Rows(156).Item("NUME_LAYER") = "TEXT_PROV_STATE"
            Data_table_mapping_layers.Rows(157).Item("NUME_LAYER") = "TEXT_RAILWAY"
            Data_table_mapping_layers.Rows(158).Item("NUME_LAYER") = "TEXT_RESERVE"
            Data_table_mapping_layers.Rows(159).Item("NUME_LAYER") = "TEXT_ROAD"
            Data_table_mapping_layers.Rows(160).Item("NUME_LAYER") = "TEXT_SMS"
            Data_table_mapping_layers.Rows(161).Item("NUME_LAYER") = "TEXT_SUBDIVISION"
            Data_table_mapping_layers.Rows(162).Item("NUME_LAYER") = "TEXT_SYMBOL"
            Data_table_mapping_layers.Rows(163).Item("NUME_LAYER") = "TEXT_TCPL"
            Data_table_mapping_layers.Rows(164).Item("NUME_LAYER") = "TEXT_TWP_RGE"
            Data_table_mapping_layers.Rows(165).Item("NUME_LAYER") = "TEXT_UTILITY"
            Data_table_mapping_layers.Rows(166).Item("NUME_LAYER") = "TEXT_WATER"
            Data_table_mapping_layers.Rows(167).Item("NUME_LAYER") = "TTEXT"
            Data_table_mapping_layers.Rows(168).Item("NUME_LAYER") = "UTILITY"
            Data_table_mapping_layers.Rows(169).Item("NUME_LAYER") = "VPORT"
            Data_table_mapping_layers.Rows(170).Item("NUME_LAYER") = "WATERBODY"
            Data_table_mapping_layers.Rows(171).Item("NUME_LAYER") = "WATERBODY_HATCH"
            Data_table_mapping_layers.Rows(172).Item("NUME_LAYER") = "WATERCOURSE"
            Data_table_mapping_layers.Rows(173).Item("NUME_LAYER") = "WOODED_AREA"
            Data_table_mapping_layers.Rows(174).Item("NUME_LAYER") = "WORKROOM"


            Data_table_mapping_layers.Rows(0).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 100-1"
            Data_table_mapping_layers.Rows(1).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 100-2"
            Data_table_mapping_layers.Rows(2).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 100-3"
            Data_table_mapping_layers.Rows(3).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 100-4"
            Data_table_mapping_layers.Rows(4).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 100-5"
            Data_table_mapping_layers.Rows(5).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 100-6"
            Data_table_mapping_layers.Rows(6).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 100-7"
            Data_table_mapping_layers.Rows(7).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1100-1"
            Data_table_mapping_layers.Rows(8).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1200-1"
            Data_table_mapping_layers.Rows(9).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1200-2"
            Data_table_mapping_layers.Rows(10).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1300-1"
            Data_table_mapping_layers.Rows(11).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1300-2"
            Data_table_mapping_layers.Rows(12).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1400-1"
            Data_table_mapping_layers.Rows(13).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1400-2"
            Data_table_mapping_layers.Rows(14).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1600-1"
            Data_table_mapping_layers.Rows(15).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 1700-1"
            Data_table_mapping_layers.Rows(16).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 200-1"
            Data_table_mapping_layers.Rows(17).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 200-2"
            Data_table_mapping_layers.Rows(18).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 200-3"
            Data_table_mapping_layers.Rows(19).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 300-1"
            Data_table_mapping_layers.Rows(20).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 300-2"
            Data_table_mapping_layers.Rows(21).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 400-1"
            Data_table_mapping_layers.Rows(22).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 400-2"
            Data_table_mapping_layers.Rows(23).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 400-3"
            Data_table_mapping_layers.Rows(24).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 500-1"
            Data_table_mapping_layers.Rows(25).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 500-2"
            Data_table_mapping_layers.Rows(26).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 500-3"
            Data_table_mapping_layers.Rows(27).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 700-1"
            Data_table_mapping_layers.Rows(28).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 700-2"
            Data_table_mapping_layers.Rows(29).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 800-1"
            Data_table_mapping_layers.Rows(30).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 800-2"
            Data_table_mapping_layers.Rows(31).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 900-1"
            Data_table_mapping_layers.Rows(32).Item("DESCRIERE_LAYER") = "TCPL Pipeline for Line No. 900-2"
            Data_table_mapping_layers.Rows(33).Item("DESCRIERE_LAYER") = "Alberta Sysytem Compressor Station"
            Data_table_mapping_layers.Rows(34).Item("DESCRIERE_LAYER") = "Alberta System Delivery Sales Station Point"
            Data_table_mapping_layers.Rows(35).Item("DESCRIERE_LAYER") = "Alberta System Future Pipeline"
            Data_table_mapping_layers.Rows(36).Item("DESCRIERE_LAYER") = "Alberta System Pipeline"
            Data_table_mapping_layers.Rows(37).Item("DESCRIERE_LAYER") = "Alberta System Proposed Pipeline and Facilities"
            Data_table_mapping_layers.Rows(38).Item("DESCRIERE_LAYER") = "Alberta System Receipt Meter Station Point"
            Data_table_mapping_layers.Rows(39).Item("DESCRIERE_LAYER") = "Boundary of Airstrip"
            Data_table_mapping_layers.Rows(40).Item("DESCRIERE_LAYER") = "Limit of ownership parcel. Includes limits Railways, Road Limits and road deviations from original township surveys"
            Data_table_mapping_layers.Rows(41).Item("DESCRIERE_LAYER") = "Line defining the limit of a District"
            Data_table_mapping_layers.Rows(42).Item("DESCRIERE_LAYER") = "Limit of federal electoral boundaries"
            Data_table_mapping_layers.Rows(43).Item("DESCRIERE_LAYER") = "Limit of a Forest Reserve (hatch colour light green)"
            Data_table_mapping_layers.Rows(44).Item("DESCRIERE_LAYER") = "Limit of Fibre optics cable easement (FOTS)"
            Data_table_mapping_layers.Rows(45).Item("DESCRIERE_LAYER") = "Limit of an indian Reserve (Hatch colour grey)"
            Data_table_mapping_layers.Rows(46).Item("DESCRIERE_LAYER") = "International Boundary"
            Data_table_mapping_layers.Rows(47).Item("DESCRIERE_LAYER") = "Limit of Surveyed lot or section"
            Data_table_mapping_layers.Rows(48).Item("DESCRIERE_LAYER") = "Limit of Unsurveyed lot or section"
            Data_table_mapping_layers.Rows(49).Item("DESCRIERE_LAYER") = "Limit of an metis Settlement"
            Data_table_mapping_layers.Rows(50).Item("DESCRIERE_LAYER") = "Limit of a Military Reserve"
            Data_table_mapping_layers.Rows(51).Item("DESCRIERE_LAYER") = "Limit of TCPL Operation Region"
            Data_table_mapping_layers.Rows(52).Item("DESCRIERE_LAYER") = "Limit of National, Provincial, Municipal or Private Park (hatch green)"
            Data_table_mapping_layers.Rows(53).Item("DESCRIERE_LAYER") = "Boundary of a Proposed subdivision"
            Data_table_mapping_layers.Rows(54).Item("DESCRIERE_LAYER") = "Provincial Boundary"
            Data_table_mapping_layers.Rows(55).Item("DESCRIERE_LAYER") = "Limit of road allowance"
            Data_table_mapping_layers.Rows(56).Item("DESCRIERE_LAYER") = "Limit of foreign right-of-way"
            Data_table_mapping_layers.Rows(57).Item("DESCRIERE_LAYER") = "State boundary"
            Data_table_mapping_layers.Rows(58).Item("DESCRIERE_LAYER") = "Limit of TCPL fee property"
            Data_table_mapping_layers.Rows(59).Item("DESCRIERE_LAYER") = "Limit of proposed TCPL easement"
            Data_table_mapping_layers.Rows(60).Item("DESCRIERE_LAYER") = "Limit of TCPL easement"
            Data_table_mapping_layers.Rows(61).Item("DESCRIERE_LAYER") = "TCPL surveyed monumented line"
            Data_table_mapping_layers.Rows(62).Item("DESCRIERE_LAYER") = "Work Room Limit"
            Data_table_mapping_layers.Rows(63).Item("DESCRIERE_LAYER") = "Limit of Township"
            Data_table_mapping_layers.Rows(64).Item("DESCRIERE_LAYER") = "Border Base of the drawing"
            Data_table_mapping_layers.Rows(65).Item("DESCRIERE_LAYER") = "Outline for large buildings or symbol for small buildings"
            Data_table_mapping_layers.Rows(66).Item("DESCRIERE_LAYER") = "Buildings that are needed but not plotted"
            Data_table_mapping_layers.Rows(67).Item("DESCRIERE_LAYER") = "Centreline of feature where limits of features are displayed"
            Data_table_mapping_layers.Rows(68).Item("DESCRIERE_LAYER") = "Outline of city Limit or symbol (hatch)"
            Data_table_mapping_layers.Rows(69).Item("DESCRIERE_LAYER") = "Polyline joining points of equal elevation"
            Data_table_mapping_layers.Rows(70).Item("DESCRIERE_LAYER") = "Symbol for bench mark, temporary bench mark or survey control markers"
            Data_table_mapping_layers.Rows(71).Item("DESCRIERE_LAYER") = "Limits of cut line through wooded areas or Centreline of cutline"
            Data_table_mapping_layers.Rows(72).Item("DESCRIERE_LAYER") = "Limit of Dam or symbol"
            Data_table_mapping_layers.Rows(73).Item("DESCRIERE_LAYER") = "Limit of beaver dam or symbol"
            Data_table_mapping_layers.Rows(74).Item("DESCRIERE_LAYER") = "Dimension Text"
            Data_table_mapping_layers.Rows(75).Item("DESCRIERE_LAYER") = "Limits of ditch or centreline of ditch"
            Data_table_mapping_layers.Rows(76).Item("DESCRIERE_LAYER") = "Spot elevation"
            Data_table_mapping_layers.Rows(77).Item("DESCRIERE_LAYER") = "Fence line"
            Data_table_mapping_layers.Rows(78).Item("DESCRIERE_LAYER") = "Symbol indicating direction of flow of drainage or pipeline"
            Data_table_mapping_layers.Rows(79).Item("DESCRIERE_LAYER") = "Foothills Pipeline Kilometre Post Symbol"
            Data_table_mapping_layers.Rows(80).Item("DESCRIERE_LAYER") = "Foothills Pipeline"
            Data_table_mapping_layers.Rows(81).Item("DESCRIERE_LAYER") = "Limit of Foothills right-of-way"
            Data_table_mapping_layers.Rows(82).Item("DESCRIERE_LAYER") = "Symbol for Foothills Sales Meter Station"
            Data_table_mapping_layers.Rows(83).Item("DESCRIERE_LAYER") = "Text describing Foothills Pipeline and Facilities"
            Data_table_mapping_layers.Rows(84).Item("DESCRIERE_LAYER") = "Symbol Foothills for Valve"
            Data_table_mapping_layers.Rows(85).Item("DESCRIERE_LAYER") = "All Foreign Pipeline Unknown"
            Data_table_mapping_layers.Rows(86).Item("DESCRIERE_LAYER") = "Fibre Optics Cable (FOTS)"
            Data_table_mapping_layers.Rows(87).Item("DESCRIERE_LAYER") = "Limit of Glacial boundary"
            Data_table_mapping_layers.Rows(88).Item("DESCRIERE_LAYER") = "Line showing ground level"
            Data_table_mapping_layers.Rows(89).Item("DESCRIERE_LAYER") = "Grid lines for Alberta"
            Data_table_mapping_layers.Rows(90).Item("DESCRIERE_LAYER") = "Grid lines for British Columbia"
            Data_table_mapping_layers.Rows(91).Item("DESCRIERE_LAYER") = "Grid for NTS"
            Data_table_mapping_layers.Rows(92).Item("DESCRIERE_LAYER") = "Grid lines for Saskatchewan"
            Data_table_mapping_layers.Rows(93).Item("DESCRIERE_LAYER") = "Grid lines"
            Data_table_mapping_layers.Rows(94).Item("DESCRIERE_LAYER") = "Highway or major road"
            Data_table_mapping_layers.Rows(95).Item("DESCRIERE_LAYER") = "Limit of Permanent Ice pack"
            Data_table_mapping_layers.Rows(96).Item("DESCRIERE_LAYER") = "Image is inserted onto a drawing"
            Data_table_mapping_layers.Rows(97).Item("DESCRIERE_LAYER") = "Symbol of kilometre post"
            Data_table_mapping_layers.Rows(98).Item("DESCRIERE_LAYER") = "All data for the Legend on the drawing"
            Data_table_mapping_layers.Rows(99).Item("DESCRIERE_LAYER") = "TransCanada Line No. 1"
            Data_table_mapping_layers.Rows(100).Item("DESCRIERE_LAYER") = "TransCanada Line No.2"
            Data_table_mapping_layers.Rows(101).Item("DESCRIERE_LAYER") = "TransCanada Line No. 3"
            Data_table_mapping_layers.Rows(102).Item("DESCRIERE_LAYER") = "TransCanada Line No. 4"
            Data_table_mapping_layers.Rows(103).Item("DESCRIERE_LAYER") = "TransCanada Line No. 5"
            Data_table_mapping_layers.Rows(104).Item("DESCRIERE_LAYER") = "TransCanada Line No. 6"
            Data_table_mapping_layers.Rows(105).Item("DESCRIERE_LAYER") = "TransCanada Line No. 7"
            Data_table_mapping_layers.Rows(106).Item("DESCRIERE_LAYER") = "Medium thickness line on the drawing"
            Data_table_mapping_layers.Rows(107).Item("DESCRIERE_LAYER") = "Thick line on the drawing"
            Data_table_mapping_layers.Rows(108).Item("DESCRIERE_LAYER") = "Thin line on the drawing"
            Data_table_mapping_layers.Rows(109).Item("DESCRIERE_LAYER") = "Holds the Logo for the map"
            Data_table_mapping_layers.Rows(110).Item("DESCRIERE_LAYER") = "Main Line Valve Symbol"
            Data_table_mapping_layers.Rows(111).Item("DESCRIERE_LAYER") = "Line defining limit of map data on printed sheet"
            Data_table_mapping_layers.Rows(112).Item("DESCRIERE_LAYER") = "North arrow"
            Data_table_mapping_layers.Rows(113).Item("DESCRIERE_LAYER") = "Layer that has been turned off which holds items that are needed on the drawing but not plotted. NOT a JUNK layer"
            Data_table_mapping_layers.Rows(114).Item("DESCRIERE_LAYER") = "Foreign Gas /Oil /Water /Unknown Pipeline"
            Data_table_mapping_layers.Rows(115).Item("DESCRIERE_LAYER") = "Pipeline carries gas"
            Data_table_mapping_layers.Rows(116).Item("DESCRIERE_LAYER") = "Pipeline carries multiple products"
            Data_table_mapping_layers.Rows(117).Item("DESCRIERE_LAYER") = "Pipeline carries oil"
            Data_table_mapping_layers.Rows(118).Item("DESCRIERE_LAYER") = "Railway line (single line or number of lines unknown)"
            Data_table_mapping_layers.Rows(119).Item("DESCRIERE_LAYER") = "Abandoned railway"
            Data_table_mapping_layers.Rows(120).Item("DESCRIERE_LAYER") = "Multiple rail lines in R/W"
            Data_table_mapping_layers.Rows(121).Item("DESCRIERE_LAYER") = "Registration Marks or Cut marks"
            Data_table_mapping_layers.Rows(122).Item("DESCRIERE_LAYER") = "Centreline for the road"
            Data_table_mapping_layers.Rows(123).Item("DESCRIERE_LAYER") = "Divided roads, Divided primary highway"
            Data_table_mapping_layers.Rows(124).Item("DESCRIERE_LAYER") = "Gravel packed roads"
            Data_table_mapping_layers.Rows(125).Item("DESCRIERE_LAYER") = "Primary Highways paved 2-4 lines ussualy"
            Data_table_mapping_layers.Rows(126).Item("DESCRIERE_LAYER") = "Paved surface"
            Data_table_mapping_layers.Rows(127).Item("DESCRIERE_LAYER") = "Secondary Highways"
            Data_table_mapping_layers.Rows(128).Item("DESCRIERE_LAYER") = "Trails, grass, roads, driveways"
            Data_table_mapping_layers.Rows(129).Item("DESCRIERE_LAYER") = "Winter Roads"
            Data_table_mapping_layers.Rows(130).Item("DESCRIERE_LAYER") = "Seismic Line"
            Data_table_mapping_layers.Rows(131).Item("DESCRIERE_LAYER") = "Scalebars"
            Data_table_mapping_layers.Rows(132).Item("DESCRIERE_LAYER") = "Text indicating an elevation of ground at a specific location"
            Data_table_mapping_layers.Rows(133).Item("DESCRIERE_LAYER") = "Layer to hold the symbols on the map"
            Data_table_mapping_layers.Rows(134).Item("DESCRIERE_LAYER") = "Layer for the symbols on the map that are needed but not plotted"
            Data_table_mapping_layers.Rows(135).Item("DESCRIERE_LAYER") = "Symbols that designate the road/highway numbers"
            Data_table_mapping_layers.Rows(136).Item("DESCRIERE_LAYER") = "TCPL Mainline Cathodic Protection Facilities"
            Data_table_mapping_layers.Rows(137).Item("DESCRIERE_LAYER") = "TCPL Mainline Compressor Station"
            Data_table_mapping_layers.Rows(138).Item("DESCRIERE_LAYER") = "TCPL Mainline Fabrication along the mainline"
            Data_table_mapping_layers.Rows(139).Item("DESCRIERE_LAYER") = "TCPL Mainline Pipeline"
            Data_table_mapping_layers.Rows(140).Item("DESCRIERE_LAYER") = "TCPL Mainline Proposed Pipeline and Facilities"
            Data_table_mapping_layers.Rows(141).Item("DESCRIERE_LAYER") = "TCPL Mainline Replacement pipe and text"
            Data_table_mapping_layers.Rows(142).Item("DESCRIERE_LAYER") = "TCPL Mainline Sales Meter Station and Taps"
            Data_table_mapping_layers.Rows(143).Item("DESCRIERE_LAYER") = "TCPL text"
            Data_table_mapping_layers.Rows(144).Item("DESCRIERE_LAYER") = "TCPL Mainline Valves"
            Data_table_mapping_layers.Rows(145).Item("DESCRIERE_LAYER") = "Text that describes the type of building"
            Data_table_mapping_layers.Rows(146).Item("DESCRIERE_LAYER") = "Text that describes Lots, Concessions, Parcels or Legal Land description"
            Data_table_mapping_layers.Rows(147).Item("DESCRIERE_LAYER") = "Text that references the City "
            Data_table_mapping_layers.Rows(148).Item("DESCRIERE_LAYER") = "Text for the contour line "
            Data_table_mapping_layers.Rows(149).Item("DESCRIERE_LAYER") = "Text that describes the facility "
            Data_table_mapping_layers.Rows(150).Item("DESCRIERE_LAYER") = "Text for federal electoral boundaries "
            Data_table_mapping_layers.Rows(151).Item("DESCRIERE_LAYER") = "General text"
            Data_table_mapping_layers.Rows(152).Item("DESCRIERE_LAYER") = "Text that describes the grid lines "
            Data_table_mapping_layers.Rows(153).Item("DESCRIERE_LAYER") = "Text for the Legend "
            Data_table_mapping_layers.Rows(154).Item("DESCRIERE_LAYER") = "Text that refers to the National, Provincial, Municipal or Private Park"
            Data_table_mapping_layers.Rows(155).Item("DESCRIERE_LAYER") = "Text that describes the pipe information "
            Data_table_mapping_layers.Rows(156).Item("DESCRIERE_LAYER") = "Text describing the Province or State "
            Data_table_mapping_layers.Rows(157).Item("DESCRIERE_LAYER") = "Text the references the railway line"
            Data_table_mapping_layers.Rows(158).Item("DESCRIERE_LAYER") = "Text describing Indian Reserve, Metis Settlement, Forest Reserve, Military Reserve"
            Data_table_mapping_layers.Rows(159).Item("DESCRIERE_LAYER") = "Text the references the road or highway name"
            Data_table_mapping_layers.Rows(160).Item("DESCRIERE_LAYER") = "Text that describes the Sales Meter Station or Tap "
            Data_table_mapping_layers.Rows(161).Item("DESCRIERE_LAYER") = "Text describing a proposed or existing Subdivision "
            Data_table_mapping_layers.Rows(162).Item("DESCRIERE_LAYER") = "Text that describes the symbol block "
            Data_table_mapping_layers.Rows(163).Item("DESCRIERE_LAYER") = "TCPL text"
            Data_table_mapping_layers.Rows(164).Item("DESCRIERE_LAYER") = "Text for the Township or Range "
            Data_table_mapping_layers.Rows(165).Item("DESCRIERE_LAYER") = "Text that describes the Utility "
            Data_table_mapping_layers.Rows(166).Item("DESCRIERE_LAYER") = "Text that refers to the River, Stream, Lake or Waterbody"
            Data_table_mapping_layers.Rows(167).Item("DESCRIERE_LAYER") = "Text"
            Data_table_mapping_layers.Rows(168).Item("DESCRIERE_LAYER") = "Foreign Utility above or below ground"
            Data_table_mapping_layers.Rows(169).Item("DESCRIERE_LAYER") = "Layer used to create a view port in paper space"
            Data_table_mapping_layers.Rows(170).Item("DESCRIERE_LAYER") = "Edge of all double line drainage features (wide river and lakes)"
            Data_table_mapping_layers.Rows(171).Item("DESCRIERE_LAYER") = "Hatch (fill) pattern for the Waterbody layer- For Flooded Areas use dashed pattern"
            Data_table_mapping_layers.Rows(172).Item("DESCRIERE_LAYER") = "Single line drainage feature (rivers, creeks, streams)"
            Data_table_mapping_layers.Rows(173).Item("DESCRIERE_LAYER") = "Limit of wooded area, outline of vegetation"
            Data_table_mapping_layers.Rows(174).Item("DESCRIERE_LAYER") = "Workroom limits"

            Data_table_mapping_layers.Rows(0).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(1).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(2).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(3).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(4).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(5).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(6).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(7).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(8).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(9).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(10).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(11).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(12).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(13).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(14).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(15).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(16).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(17).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(18).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(19).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(20).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(21).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(22).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(23).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(24).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(25).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(26).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(27).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(28).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(29).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(30).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(31).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(32).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(33).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(34).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(35).Item("CULOARE") = 10
            Data_table_mapping_layers.Rows(36).Item("CULOARE") = 90
            Data_table_mapping_layers.Rows(37).Item("CULOARE") = 10
            Data_table_mapping_layers.Rows(38).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(39).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(40).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(41).Item("CULOARE") = 254
            Data_table_mapping_layers.Rows(42).Item("CULOARE") = 191
            Data_table_mapping_layers.Rows(43).Item("CULOARE") = 61
            Data_table_mapping_layers.Rows(44).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(45).Item("CULOARE") = 254
            Data_table_mapping_layers.Rows(46).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(47).Item("CULOARE") = 253
            Data_table_mapping_layers.Rows(48).Item("CULOARE") = 254
            Data_table_mapping_layers.Rows(49).Item("CULOARE") = 253
            Data_table_mapping_layers.Rows(50).Item("CULOARE") = 252
            Data_table_mapping_layers.Rows(51).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(52).Item("CULOARE") = 71
            Data_table_mapping_layers.Rows(53).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(54).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(55).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(56).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(57).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(58).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(59).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(60).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(61).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(62).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(63).Item("CULOARE") = 254
            Data_table_mapping_layers.Rows(64).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(65).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(66).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(67).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(68).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(69).Item("CULOARE") = 31
            Data_table_mapping_layers.Rows(70).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(71).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(72).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(73).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(74).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(75).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(76).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(77).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(78).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(79).Item("CULOARE") = 200
            Data_table_mapping_layers.Rows(80).Item("CULOARE") = 200
            Data_table_mapping_layers.Rows(81).Item("CULOARE") = 200
            Data_table_mapping_layers.Rows(82).Item("CULOARE") = 200
            Data_table_mapping_layers.Rows(83).Item("CULOARE") = 200
            Data_table_mapping_layers.Rows(84).Item("CULOARE") = 200
            Data_table_mapping_layers.Rows(85).Item("CULOARE") = 200
            Data_table_mapping_layers.Rows(86).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(87).Item("CULOARE") = 30
            Data_table_mapping_layers.Rows(88).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(89).Item("CULOARE") = 251
            Data_table_mapping_layers.Rows(90).Item("CULOARE") = 251
            Data_table_mapping_layers.Rows(91).Item("CULOARE") = 251
            Data_table_mapping_layers.Rows(92).Item("CULOARE") = 251
            Data_table_mapping_layers.Rows(93).Item("CULOARE") = 251
            Data_table_mapping_layers.Rows(94).Item("CULOARE") = 34
            Data_table_mapping_layers.Rows(95).Item("CULOARE") = 130
            Data_table_mapping_layers.Rows(96).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(97).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(98).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(99).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(100).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(101).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(102).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(103).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(104).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(105).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(106).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(107).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(108).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(109).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(110).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(111).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(112).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(113).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(114).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(115).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(116).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(117).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(118).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(119).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(120).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(121).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(122).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(123).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(124).Item("CULOARE") = 30
            Data_table_mapping_layers.Rows(125).Item("CULOARE") = 34
            Data_table_mapping_layers.Rows(126).Item("CULOARE") = 34
            Data_table_mapping_layers.Rows(127).Item("CULOARE") = 34
            Data_table_mapping_layers.Rows(128).Item("CULOARE") = 30
            Data_table_mapping_layers.Rows(129).Item("CULOARE") = 30
            Data_table_mapping_layers.Rows(130).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(131).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(132).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(133).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(134).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(135).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(136).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(137).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(138).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(139).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(140).Item("CULOARE") = 10
            Data_table_mapping_layers.Rows(141).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(142).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(143).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(144).Item("CULOARE") = 160
            Data_table_mapping_layers.Rows(145).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(146).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(147).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(148).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(149).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(150).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(151).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(152).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(153).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(154).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(155).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(156).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(157).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(158).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(159).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(160).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(161).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(162).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(163).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(164).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(165).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(166).Item("CULOARE") = 150
            Data_table_mapping_layers.Rows(167).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(168).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(169).Item("CULOARE") = 7
            Data_table_mapping_layers.Rows(170).Item("CULOARE") = 150
            Data_table_mapping_layers.Rows(171).Item("CULOARE") = 131
            Data_table_mapping_layers.Rows(172).Item("CULOARE") = 150
            Data_table_mapping_layers.Rows(173).Item("CULOARE") = 81
            Data_table_mapping_layers.Rows(174).Item("CULOARE") = 7

            Data_table_mapping_layers.Rows(0).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(1).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(2).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(3).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(4).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(5).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(6).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(7).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(8).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(9).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(10).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(11).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(12).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(13).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(14).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(15).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(16).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(17).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(18).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(19).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(20).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(21).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(22).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(23).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(24).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(25).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(26).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(27).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(28).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(29).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(30).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(31).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(32).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(33).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(34).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(35).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(36).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(37).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(38).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(39).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(40).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(41).Item("LINEWEIGHT") = LineWeight.LineWeight100
            Data_table_mapping_layers.Rows(42).Item("LINEWEIGHT") = LineWeight.LineWeight070
            Data_table_mapping_layers.Rows(43).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(44).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(45).Item("LINEWEIGHT") = LineWeight.LineWeight100
            Data_table_mapping_layers.Rows(46).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(47).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(48).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(49).Item("LINEWEIGHT") = LineWeight.LineWeight030
            Data_table_mapping_layers.Rows(50).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(51).Item("LINEWEIGHT") = LineWeight.LineWeight140
            Data_table_mapping_layers.Rows(52).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(53).Item("LINEWEIGHT") = LineWeight.LineWeight030
            Data_table_mapping_layers.Rows(54).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(55).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(56).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(57).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(58).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(59).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(60).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(61).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(62).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(63).Item("LINEWEIGHT") = LineWeight.LineWeight100
            Data_table_mapping_layers.Rows(64).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(65).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(66).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(67).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(68).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(69).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(70).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(71).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(72).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(73).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(74).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(75).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(76).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(77).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(78).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(79).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(80).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(81).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(82).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(83).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(84).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(85).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(86).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(87).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(88).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(89).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(90).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(91).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(92).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(93).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(94).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(95).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(96).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(97).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(98).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(99).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(100).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(101).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(102).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(103).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(104).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(105).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(106).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(107).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_mapping_layers.Rows(108).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(109).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(110).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(111).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(112).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(113).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(114).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(115).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(116).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(117).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(118).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(119).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(120).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(121).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(122).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(123).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(124).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(125).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(126).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(127).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(128).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(129).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(130).Item("LINEWEIGHT") = LineWeight.LineWeight020
            Data_table_mapping_layers.Rows(131).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(132).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(133).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(134).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(135).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(136).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(137).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(138).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(139).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(140).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(141).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(142).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(143).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(144).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(145).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(146).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(147).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(148).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(149).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(150).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(151).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(152).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(153).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(154).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(155).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(156).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(157).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(158).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(159).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(160).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(161).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(162).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(163).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(164).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(165).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(166).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(167).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(168).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(169).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(170).Item("LINEWEIGHT") = LineWeight.LineWeight035
            Data_table_mapping_layers.Rows(171).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(172).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(173).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_mapping_layers.Rows(174).Item("LINEWEIGHT") = LineWeight.LineWeight025

            Data_table_mapping_layers.Rows(0).Item("PLOT") = True
            Data_table_mapping_layers.Rows(1).Item("PLOT") = True
            Data_table_mapping_layers.Rows(2).Item("PLOT") = True
            Data_table_mapping_layers.Rows(3).Item("PLOT") = True
            Data_table_mapping_layers.Rows(4).Item("PLOT") = True
            Data_table_mapping_layers.Rows(5).Item("PLOT") = True
            Data_table_mapping_layers.Rows(6).Item("PLOT") = True
            Data_table_mapping_layers.Rows(7).Item("PLOT") = True
            Data_table_mapping_layers.Rows(8).Item("PLOT") = True
            Data_table_mapping_layers.Rows(9).Item("PLOT") = True
            Data_table_mapping_layers.Rows(10).Item("PLOT") = True
            Data_table_mapping_layers.Rows(11).Item("PLOT") = True
            Data_table_mapping_layers.Rows(12).Item("PLOT") = True
            Data_table_mapping_layers.Rows(13).Item("PLOT") = True
            Data_table_mapping_layers.Rows(14).Item("PLOT") = True
            Data_table_mapping_layers.Rows(15).Item("PLOT") = True
            Data_table_mapping_layers.Rows(16).Item("PLOT") = True
            Data_table_mapping_layers.Rows(17).Item("PLOT") = True
            Data_table_mapping_layers.Rows(18).Item("PLOT") = True
            Data_table_mapping_layers.Rows(19).Item("PLOT") = True
            Data_table_mapping_layers.Rows(20).Item("PLOT") = True
            Data_table_mapping_layers.Rows(21).Item("PLOT") = True
            Data_table_mapping_layers.Rows(22).Item("PLOT") = True
            Data_table_mapping_layers.Rows(23).Item("PLOT") = True
            Data_table_mapping_layers.Rows(24).Item("PLOT") = True
            Data_table_mapping_layers.Rows(25).Item("PLOT") = True
            Data_table_mapping_layers.Rows(26).Item("PLOT") = True
            Data_table_mapping_layers.Rows(27).Item("PLOT") = True
            Data_table_mapping_layers.Rows(28).Item("PLOT") = True
            Data_table_mapping_layers.Rows(29).Item("PLOT") = True
            Data_table_mapping_layers.Rows(30).Item("PLOT") = True
            Data_table_mapping_layers.Rows(31).Item("PLOT") = True
            Data_table_mapping_layers.Rows(32).Item("PLOT") = True
            Data_table_mapping_layers.Rows(33).Item("PLOT") = True
            Data_table_mapping_layers.Rows(34).Item("PLOT") = True
            Data_table_mapping_layers.Rows(35).Item("PLOT") = True
            Data_table_mapping_layers.Rows(36).Item("PLOT") = True
            Data_table_mapping_layers.Rows(37).Item("PLOT") = True
            Data_table_mapping_layers.Rows(38).Item("PLOT") = True
            Data_table_mapping_layers.Rows(39).Item("PLOT") = True
            Data_table_mapping_layers.Rows(40).Item("PLOT") = True
            Data_table_mapping_layers.Rows(41).Item("PLOT") = True
            Data_table_mapping_layers.Rows(42).Item("PLOT") = True
            Data_table_mapping_layers.Rows(43).Item("PLOT") = True
            Data_table_mapping_layers.Rows(44).Item("PLOT") = True
            Data_table_mapping_layers.Rows(45).Item("PLOT") = True
            Data_table_mapping_layers.Rows(46).Item("PLOT") = True
            Data_table_mapping_layers.Rows(47).Item("PLOT") = True
            Data_table_mapping_layers.Rows(48).Item("PLOT") = True
            Data_table_mapping_layers.Rows(49).Item("PLOT") = True
            Data_table_mapping_layers.Rows(50).Item("PLOT") = True
            Data_table_mapping_layers.Rows(51).Item("PLOT") = True
            Data_table_mapping_layers.Rows(52).Item("PLOT") = True
            Data_table_mapping_layers.Rows(53).Item("PLOT") = True
            Data_table_mapping_layers.Rows(54).Item("PLOT") = True
            Data_table_mapping_layers.Rows(55).Item("PLOT") = True
            Data_table_mapping_layers.Rows(56).Item("PLOT") = True
            Data_table_mapping_layers.Rows(57).Item("PLOT") = True
            Data_table_mapping_layers.Rows(58).Item("PLOT") = True
            Data_table_mapping_layers.Rows(59).Item("PLOT") = True
            Data_table_mapping_layers.Rows(60).Item("PLOT") = True
            Data_table_mapping_layers.Rows(61).Item("PLOT") = True
            Data_table_mapping_layers.Rows(62).Item("PLOT") = True
            Data_table_mapping_layers.Rows(63).Item("PLOT") = True
            Data_table_mapping_layers.Rows(64).Item("PLOT") = True
            Data_table_mapping_layers.Rows(65).Item("PLOT") = True
            Data_table_mapping_layers.Rows(66).Item("PLOT") = False
            Data_table_mapping_layers.Rows(67).Item("PLOT") = True
            Data_table_mapping_layers.Rows(68).Item("PLOT") = True
            Data_table_mapping_layers.Rows(69).Item("PLOT") = True
            Data_table_mapping_layers.Rows(70).Item("PLOT") = True
            Data_table_mapping_layers.Rows(71).Item("PLOT") = True
            Data_table_mapping_layers.Rows(72).Item("PLOT") = True
            Data_table_mapping_layers.Rows(73).Item("PLOT") = True
            Data_table_mapping_layers.Rows(74).Item("PLOT") = True
            Data_table_mapping_layers.Rows(75).Item("PLOT") = True
            Data_table_mapping_layers.Rows(76).Item("PLOT") = True
            Data_table_mapping_layers.Rows(77).Item("PLOT") = True
            Data_table_mapping_layers.Rows(78).Item("PLOT") = True
            Data_table_mapping_layers.Rows(79).Item("PLOT") = True
            Data_table_mapping_layers.Rows(80).Item("PLOT") = True
            Data_table_mapping_layers.Rows(81).Item("PLOT") = True
            Data_table_mapping_layers.Rows(82).Item("PLOT") = True
            Data_table_mapping_layers.Rows(83).Item("PLOT") = True
            Data_table_mapping_layers.Rows(84).Item("PLOT") = True
            Data_table_mapping_layers.Rows(85).Item("PLOT") = True
            Data_table_mapping_layers.Rows(86).Item("PLOT") = True
            Data_table_mapping_layers.Rows(87).Item("PLOT") = True
            Data_table_mapping_layers.Rows(88).Item("PLOT") = True
            Data_table_mapping_layers.Rows(89).Item("PLOT") = True
            Data_table_mapping_layers.Rows(90).Item("PLOT") = True
            Data_table_mapping_layers.Rows(91).Item("PLOT") = True
            Data_table_mapping_layers.Rows(92).Item("PLOT") = True
            Data_table_mapping_layers.Rows(93).Item("PLOT") = True
            Data_table_mapping_layers.Rows(94).Item("PLOT") = True
            Data_table_mapping_layers.Rows(95).Item("PLOT") = True
            Data_table_mapping_layers.Rows(96).Item("PLOT") = True
            Data_table_mapping_layers.Rows(97).Item("PLOT") = True
            Data_table_mapping_layers.Rows(98).Item("PLOT") = True
            Data_table_mapping_layers.Rows(99).Item("PLOT") = True
            Data_table_mapping_layers.Rows(100).Item("PLOT") = True
            Data_table_mapping_layers.Rows(101).Item("PLOT") = True
            Data_table_mapping_layers.Rows(102).Item("PLOT") = True
            Data_table_mapping_layers.Rows(103).Item("PLOT") = True
            Data_table_mapping_layers.Rows(104).Item("PLOT") = True
            Data_table_mapping_layers.Rows(105).Item("PLOT") = True
            Data_table_mapping_layers.Rows(106).Item("PLOT") = True
            Data_table_mapping_layers.Rows(107).Item("PLOT") = True
            Data_table_mapping_layers.Rows(108).Item("PLOT") = True
            Data_table_mapping_layers.Rows(109).Item("PLOT") = True
            Data_table_mapping_layers.Rows(110).Item("PLOT") = True
            Data_table_mapping_layers.Rows(111).Item("PLOT") = True
            Data_table_mapping_layers.Rows(112).Item("PLOT") = True
            Data_table_mapping_layers.Rows(113).Item("PLOT") = False
            Data_table_mapping_layers.Rows(114).Item("PLOT") = True
            Data_table_mapping_layers.Rows(115).Item("PLOT") = True
            Data_table_mapping_layers.Rows(116).Item("PLOT") = True
            Data_table_mapping_layers.Rows(117).Item("PLOT") = True
            Data_table_mapping_layers.Rows(118).Item("PLOT") = True
            Data_table_mapping_layers.Rows(119).Item("PLOT") = True
            Data_table_mapping_layers.Rows(120).Item("PLOT") = True
            Data_table_mapping_layers.Rows(121).Item("PLOT") = True
            Data_table_mapping_layers.Rows(122).Item("PLOT") = True
            Data_table_mapping_layers.Rows(123).Item("PLOT") = True
            Data_table_mapping_layers.Rows(124).Item("PLOT") = True
            Data_table_mapping_layers.Rows(125).Item("PLOT") = True
            Data_table_mapping_layers.Rows(126).Item("PLOT") = True
            Data_table_mapping_layers.Rows(127).Item("PLOT") = True
            Data_table_mapping_layers.Rows(128).Item("PLOT") = True
            Data_table_mapping_layers.Rows(129).Item("PLOT") = True
            Data_table_mapping_layers.Rows(130).Item("PLOT") = True
            Data_table_mapping_layers.Rows(131).Item("PLOT") = True
            Data_table_mapping_layers.Rows(132).Item("PLOT") = True
            Data_table_mapping_layers.Rows(133).Item("PLOT") = True
            Data_table_mapping_layers.Rows(134).Item("PLOT") = True
            Data_table_mapping_layers.Rows(135).Item("PLOT") = True
            Data_table_mapping_layers.Rows(136).Item("PLOT") = True
            Data_table_mapping_layers.Rows(137).Item("PLOT") = True
            Data_table_mapping_layers.Rows(138).Item("PLOT") = True
            Data_table_mapping_layers.Rows(139).Item("PLOT") = True
            Data_table_mapping_layers.Rows(140).Item("PLOT") = True
            Data_table_mapping_layers.Rows(141).Item("PLOT") = True
            Data_table_mapping_layers.Rows(142).Item("PLOT") = True
            Data_table_mapping_layers.Rows(143).Item("PLOT") = True
            Data_table_mapping_layers.Rows(144).Item("PLOT") = True
            Data_table_mapping_layers.Rows(145).Item("PLOT") = True
            Data_table_mapping_layers.Rows(146).Item("PLOT") = True
            Data_table_mapping_layers.Rows(147).Item("PLOT") = True
            Data_table_mapping_layers.Rows(148).Item("PLOT") = True
            Data_table_mapping_layers.Rows(149).Item("PLOT") = True
            Data_table_mapping_layers.Rows(150).Item("PLOT") = True
            Data_table_mapping_layers.Rows(151).Item("PLOT") = True
            Data_table_mapping_layers.Rows(152).Item("PLOT") = True
            Data_table_mapping_layers.Rows(153).Item("PLOT") = True
            Data_table_mapping_layers.Rows(154).Item("PLOT") = True
            Data_table_mapping_layers.Rows(155).Item("PLOT") = True
            Data_table_mapping_layers.Rows(156).Item("PLOT") = True
            Data_table_mapping_layers.Rows(157).Item("PLOT") = True
            Data_table_mapping_layers.Rows(158).Item("PLOT") = True
            Data_table_mapping_layers.Rows(159).Item("PLOT") = True
            Data_table_mapping_layers.Rows(160).Item("PLOT") = True
            Data_table_mapping_layers.Rows(161).Item("PLOT") = True
            Data_table_mapping_layers.Rows(162).Item("PLOT") = True
            Data_table_mapping_layers.Rows(163).Item("PLOT") = True
            Data_table_mapping_layers.Rows(164).Item("PLOT") = True
            Data_table_mapping_layers.Rows(165).Item("PLOT") = True
            Data_table_mapping_layers.Rows(166).Item("PLOT") = True
            Data_table_mapping_layers.Rows(167).Item("PLOT") = True
            Data_table_mapping_layers.Rows(168).Item("PLOT") = True
            Data_table_mapping_layers.Rows(169).Item("PLOT") = False
            Data_table_mapping_layers.Rows(170).Item("PLOT") = True
            Data_table_mapping_layers.Rows(171).Item("PLOT") = True
            Data_table_mapping_layers.Rows(172).Item("PLOT") = True
            Data_table_mapping_layers.Rows(173).Item("PLOT") = True
            Data_table_mapping_layers.Rows(174).Item("PLOT") = True


            Data_table_mapping_layers.Rows(0).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(1).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(2).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(3).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(4).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(5).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(6).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(7).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(8).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(9).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(10).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(11).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(12).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(13).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(14).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(15).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(16).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(17).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(18).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(19).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(20).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(21).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(22).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(23).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(24).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(25).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(26).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(27).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(28).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(29).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(30).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(31).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(32).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(33).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(34).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(35).Item("LINETYPE") = "TCDASHED"
            Data_table_mapping_layers.Rows(36).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(37).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(38).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(39).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(40).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(41).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(42).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(43).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(44).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(45).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(46).Item("LINETYPE") = "TCPHANTOM2"
            Data_table_mapping_layers.Rows(47).Item("LINETYPE") = "TCPHANTOMX2"
            Data_table_mapping_layers.Rows(48).Item("LINETYPE") = "TCPHANTOMX2"
            Data_table_mapping_layers.Rows(49).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(50).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(51).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(52).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(53).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(54).Item("LINETYPE") = "TCCENTER3"
            Data_table_mapping_layers.Rows(55).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(56).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(57).Item("LINETYPE") = "TCCENTER3"
            Data_table_mapping_layers.Rows(58).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(59).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(60).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(61).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(62).Item("LINETYPE") = "TCDOT4"
            Data_table_mapping_layers.Rows(63).Item("LINETYPE") = "TCPHANTOMX2"
            Data_table_mapping_layers.Rows(64).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(65).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(66).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(67).Item("LINETYPE") = "TCCENTER"
            Data_table_mapping_layers.Rows(68).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(69).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(70).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(71).Item("LINETYPE") = "TCDOT4"
            Data_table_mapping_layers.Rows(72).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(73).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(74).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(75).Item("LINETYPE") = "TCCENTER"
            Data_table_mapping_layers.Rows(76).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(77).Item("LINETYPE") = "TC_FAC_FENCE"
            Data_table_mapping_layers.Rows(78).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(79).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(80).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(81).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(82).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(83).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(84).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(85).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(86).Item("LINETYPE") = "TC_UG_TEL"
            Data_table_mapping_layers.Rows(87).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(88).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(89).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(90).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(91).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(92).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(93).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(94).Item("LINETYPE") = "TCCENTER"
            Data_table_mapping_layers.Rows(95).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(96).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(97).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(98).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(99).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(100).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(101).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(102).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(103).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(104).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(105).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(106).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(107).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(108).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(109).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(110).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(111).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(112).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(113).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(114).Item("LINETYPE") = "TC_FOREIGN_PIPE"
            Data_table_mapping_layers.Rows(115).Item("LINETYPE") = "TC_GAS_LINE"
            Data_table_mapping_layers.Rows(116).Item("LINETYPE") = "TC_FOREIGN_PIPE"
            Data_table_mapping_layers.Rows(117).Item("LINETYPE") = "TC_OIL_LINE"
            Data_table_mapping_layers.Rows(118).Item("LINETYPE") = "TC_TRACK_S"
            Data_table_mapping_layers.Rows(119).Item("LINETYPE") = "TC_TRACK_A"
            Data_table_mapping_layers.Rows(120).Item("LINETYPE") = "TC_TRACK_M"
            Data_table_mapping_layers.Rows(121).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(122).Item("LINETYPE") = "TCCENTER"
            Data_table_mapping_layers.Rows(123).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(124).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(125).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(126).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(127).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(128).Item("LINETYPE") = "TCDASHED"
            Data_table_mapping_layers.Rows(129).Item("LINETYPE") = "TCDASHED"
            Data_table_mapping_layers.Rows(130).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(131).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(132).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(133).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(134).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(135).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(136).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(137).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(138).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(139).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(140).Item("LINETYPE") = "TCDASH3"
            Data_table_mapping_layers.Rows(141).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(142).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(143).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(144).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(145).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(146).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(147).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(148).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(149).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(150).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(151).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(152).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(153).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(154).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(155).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(156).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(157).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(158).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(159).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(160).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(161).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(162).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(163).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(164).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(165).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(166).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(167).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(168).Item("LINETYPE") = "TC_TELEPHONE"
            Data_table_mapping_layers.Rows(169).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(170).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(171).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(172).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(173).Item("LINETYPE") = "Continuous"
            Data_table_mapping_layers.Rows(174).Item("LINETYPE") = "TCDASHED"


            Data_table_extra_layers = New System.Data.DataTable
            Data_table_extra_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_extra_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_extra_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_extra_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_extra_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_extra_layers.Columns.Add("LINETYPE", GetType(String))


            For i = 0 To 6
                Data_table_extra_layers.Rows.Add()
            Next


            Data_table_extra_layers.Rows(0).Item("NUME_LAYER") = "MARSH"
            Data_table_extra_layers.Rows(1).Item("NUME_LAYER") = "PVEG"
            Data_table_extra_layers.Rows(2).Item("NUME_LAYER") = "ROAD_BDRY"
            Data_table_extra_layers.Rows(3).Item("NUME_LAYER") = "TEMP_WORK_SPACE"
            Data_table_extra_layers.Rows(4).Item("NUME_LAYER") = "TEXT_WS"
            Data_table_extra_layers.Rows(5).Item("NUME_LAYER") = "TRAIL"
            Data_table_extra_layers.Rows(6).Item("NUME_LAYER") = "PCENTRE"

            Data_table_extra_layers.Rows(0).Item("DESCRIERE_LAYER") = "Low wet area, marsh, swamp"
            Data_table_extra_layers.Rows(1).Item("DESCRIERE_LAYER") = "Bushline"
            Data_table_extra_layers.Rows(2).Item("DESCRIERE_LAYER") = "Road Boundary"
            Data_table_extra_layers.Rows(3).Item("DESCRIERE_LAYER") = "Line describing the limits of the temporary work space on construction projects"
            Data_table_extra_layers.Rows(4).Item("DESCRIERE_LAYER") = "Text for temporary work space size and length"
            Data_table_extra_layers.Rows(5).Item("DESCRIERE_LAYER") = "A trail passable only by people on foot, bicycle or ATV's"
            Data_table_extra_layers.Rows(6).Item("DESCRIERE_LAYER") = "Pipe centerline on the graph profile (usually off in the viewport)"

            Data_table_extra_layers.Rows(0).Item("CULOARE") = 131
            Data_table_extra_layers.Rows(1).Item("CULOARE") = 3
            Data_table_extra_layers.Rows(2).Item("CULOARE") = 7
            Data_table_extra_layers.Rows(3).Item("CULOARE") = 7
            Data_table_extra_layers.Rows(4).Item("CULOARE") = 7
            Data_table_extra_layers.Rows(5).Item("CULOARE") = 30
            Data_table_extra_layers.Rows(6).Item("CULOARE") = 30

            Data_table_extra_layers.Rows(0).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_extra_layers.Rows(1).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_extra_layers.Rows(2).Item("LINEWEIGHT") = LineWeight.LineWeight050
            Data_table_extra_layers.Rows(3).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_extra_layers.Rows(4).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_extra_layers.Rows(5).Item("LINEWEIGHT") = LineWeight.LineWeight025
            Data_table_extra_layers.Rows(6).Item("LINEWEIGHT") = LineWeight.LineWeight025

            Data_table_extra_layers.Rows(0).Item("PLOT") = True
            Data_table_extra_layers.Rows(1).Item("PLOT") = True
            Data_table_extra_layers.Rows(2).Item("PLOT") = True
            Data_table_extra_layers.Rows(3).Item("PLOT") = True
            Data_table_extra_layers.Rows(4).Item("PLOT") = True
            Data_table_extra_layers.Rows(5).Item("PLOT") = True
            Data_table_extra_layers.Rows(6).Item("PLOT") = True

            Data_table_extra_layers.Rows(0).Item("LINETYPE") = "Continuous"
            Data_table_extra_layers.Rows(1).Item("LINETYPE") = "Continuous"
            Data_table_extra_layers.Rows(2).Item("LINETYPE") = "Continuous"
            Data_table_extra_layers.Rows(3).Item("LINETYPE") = "TCDASHED"
            Data_table_extra_layers.Rows(4).Item("LINETYPE") = "Continuous"
            Data_table_extra_layers.Rows(5).Item("LINETYPE") = "TCDASHED"
            Data_table_extra_layers.Rows(6).Item("LINETYPE") = "TCCENTER"



            Data_table_my_list_layers = New System.Data.DataTable
            Data_table_my_list_layers.Columns.Add("NUME_LAYER", GetType(String))
            Data_table_my_list_layers.Columns.Add("DESCRIERE_LAYER", GetType(String))
            Data_table_my_list_layers.Columns.Add("CULOARE", GetType(Integer))
            Data_table_my_list_layers.Columns.Add("LINEWEIGHT", GetType(LineWeight))
            Data_table_my_list_layers.Columns.Add("PLOT", GetType(Boolean))
            Data_table_my_list_layers.Columns.Add("LINETYPE", GetType(String))

           
            ListBox_LAYERING_name.Visible = False
            ListBox_my_list.Visible = False
            TextBox_description.Visible = False
            Button_load_from_my_list.Visible = False
            Button_CREATE_UPDATE_MY_LIST.Visible = False
            Button_CREATE_UPDATE_MY_LIST.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = False
            Button_Read_my_list.Visible = False
            Button_LOAD_1_LAYER.Visible = False
            CheckBox_my_list.Visible = False
            Me.Width = 752
        Catch ex As Exception
            Application.SetSystemVariable("FILEDIA", 1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Incarca_toate_linetypes()
        Try

            ascunde_butoanele()
            Creaza_Dim_style_AND_text_style_ROMANS()

            afiseaza_butoanele()
            SET_FILEDIA_TO_1()
        Catch ex As Exception
            afiseaza_butoanele()
            Application.SetSystemVariable("FILEDIA", 1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_LOAD_ONLY_LTYPES_Click(sender As Object, e As EventArgs) Handles Button_LOAD_ONLY_LTYPES.Click
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            ascunde_butoanele()
            Creaza_Dim_style_AND_text_style_ROMANS()
            afiseaza_butoanele()
            SET_FILEDIA_TO_1()
        Catch ex As Exception
            afiseaza_butoanele()
            Application.SetSystemVariable("FILEDIA", 1)
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ListBox_LAYERING_name_Click(sender As Object, e As EventArgs) Handles ListBox_LAYERING_name.Click
        Try
            Dim Curent_index As Integer = ListBox_LAYERING_name.SelectedIndex
            If Not Curent_index = -1 Then
                TextBox_description.Visible = True



                If CheckBox_my_list.Checked = True Then
                    If ListBox_LAYERING_name.Items.Count > 0 Then
                        If Data_table_general_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                            If ListBox_my_list.Items.Contains(Data_table_general_layers.Rows(Curent_index).Item("NUME_LAYER")) = False Then
                                Data_table_my_list_layers.Rows.Add()
                                Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                                Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Data_table_general_layers.Rows(Curent_index).Item("NUME_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Data_table_general_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = Data_table_general_layers.Rows(Curent_index).Item("CULOARE")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = Data_table_general_layers.Rows(Curent_index).Item("LINEWEIGHT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = Data_table_general_layers.Rows(Curent_index).Item("PLOT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Data_table_general_layers.Rows(Curent_index).Item("LINETYPE")
                                ListBox_my_list.Items.Add(Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER"))
                            End If
                        End If
                        If Data_table_civil_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                            If ListBox_my_list.Items.Contains(Data_table_civil_layers.Rows(Curent_index).Item("NUME_LAYER")) = False Then
                                Data_table_my_list_layers.Rows.Add()
                                Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                                Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Data_table_civil_layers.Rows(Curent_index).Item("NUME_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Data_table_civil_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = Data_table_civil_layers.Rows(Curent_index).Item("CULOARE")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = Data_table_civil_layers.Rows(Curent_index).Item("LINEWEIGHT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = Data_table_civil_layers.Rows(Curent_index).Item("PLOT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Data_table_civil_layers.Rows(Curent_index).Item("LINETYPE")
                                ListBox_my_list.Items.Add(Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER"))
                            End If
                        End If

                        If Data_table_electrical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                            If ListBox_my_list.Items.Contains(Data_table_electrical_layers.Rows(Curent_index).Item("NUME_LAYER")) = False Then
                                Data_table_my_list_layers.Rows.Add()
                                Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                                Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Data_table_electrical_layers.Rows(Curent_index).Item("NUME_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Data_table_electrical_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = Data_table_electrical_layers.Rows(Curent_index).Item("CULOARE")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = Data_table_electrical_layers.Rows(Curent_index).Item("LINEWEIGHT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = Data_table_electrical_layers.Rows(Curent_index).Item("PLOT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Data_table_electrical_layers.Rows(Curent_index).Item("LINETYPE")
                                ListBox_my_list.Items.Add(Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER"))
                            End If
                        End If
                        If Data_table_mechanical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                            If ListBox_my_list.Items.Contains(Data_table_mechanical_layers.Rows(Curent_index).Item("NUME_LAYER")) = False Then
                                Data_table_my_list_layers.Rows.Add()
                                Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                                Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Data_table_mechanical_layers.Rows(Curent_index).Item("NUME_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Data_table_mechanical_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = Data_table_mechanical_layers.Rows(Curent_index).Item("CULOARE")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = Data_table_mechanical_layers.Rows(Curent_index).Item("LINEWEIGHT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = Data_table_mechanical_layers.Rows(Curent_index).Item("PLOT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Data_table_mechanical_layers.Rows(Curent_index).Item("LINETYPE")
                                ListBox_my_list.Items.Add(Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER"))
                            End If
                        End If
                        If Data_table_pipeline_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                            If ListBox_my_list.Items.Contains(Data_table_mechanical_layers.Rows(Curent_index).Item("NUME_LAYER")) = False Then
                                Data_table_my_list_layers.Rows.Add()
                                Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                                Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Data_table_pipeline_layers.Rows(Curent_index).Item("NUME_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Data_table_pipeline_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = Data_table_pipeline_layers.Rows(Curent_index).Item("CULOARE")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = Data_table_pipeline_layers.Rows(Curent_index).Item("LINEWEIGHT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = Data_table_pipeline_layers.Rows(Curent_index).Item("PLOT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Data_table_pipeline_layers.Rows(Curent_index).Item("LINETYPE")
                                ListBox_my_list.Items.Add(Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER"))
                            End If
                        End If
                        If Data_table_mapping_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                            If ListBox_my_list.Items.Contains(Data_table_mapping_layers.Rows(Curent_index).Item("NUME_LAYER")) = False Then
                                Data_table_my_list_layers.Rows.Add()
                                Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                                Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Data_table_mapping_layers.Rows(Curent_index).Item("NUME_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Data_table_mapping_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = Data_table_mapping_layers.Rows(Curent_index).Item("CULOARE")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = Data_table_mapping_layers.Rows(Curent_index).Item("LINEWEIGHT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = Data_table_mapping_layers.Rows(Curent_index).Item("PLOT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Data_table_mapping_layers.Rows(Curent_index).Item("LINETYPE")
                                ListBox_my_list.Items.Add(Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER"))
                            End If
                        End If
                        If Data_table_extra_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                            If ListBox_my_list.Items.Contains(Data_table_extra_layers.Rows(Curent_index).Item("NUME_LAYER")) = False Then
                                Data_table_my_list_layers.Rows.Add()
                                Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                                Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Data_table_extra_layers.Rows(Curent_index).Item("NUME_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Data_table_extra_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = Data_table_extra_layers.Rows(Curent_index).Item("CULOARE")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = Data_table_extra_layers.Rows(Curent_index).Item("LINEWEIGHT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = Data_table_extra_layers.Rows(Curent_index).Item("PLOT")
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Data_table_extra_layers.Rows(Curent_index).Item("LINETYPE")
                                ListBox_my_list.Items.Add(Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER"))
                            End If
                        End If



                       

                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

   

    Private Sub Button_LOAD_grup_Click(sender As Object, e As EventArgs) Handles Button_LOAD_GRUP_LAYERS.Click
        Try
            Incarca_toate_linetypes()
            If ListBox_LAYERING_name.Visible = True Then

                ascunde_butoanele()
                If ListBox_LAYERING_name.Items.Count > 0 Then
                    If Data_table_general_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                        For i = 0 To Data_table_general_layers.Rows.Count - 1
                            Dim Ltype As String = Data_table_general_layers.Rows(i).Item("LINETYPE")
                            Dim Nume1 As String = Data_table_general_layers.Rows(i).Item("NUME_LAYER")
                            Dim Descr As String = Data_table_general_layers.Rows(i).Item("DESCRIERE_LAYER")
                            Dim Culoare As String = Data_table_general_layers.Rows(i).Item("CULOARE")
                            Dim LineW As LineWeight = Data_table_general_layers.Rows(i).Item("LINEWEIGHT")
                            Dim Plot1 As Boolean = Data_table_general_layers.Rows(i).Item("PLOT")
                            Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                        Next
                    End If

                    If Data_table_civil_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                        For i = 0 To Data_table_civil_layers.Rows.Count - 1
                            Dim Ltype As String = Data_table_civil_layers.Rows(i).Item("LINETYPE")
                            Dim Nume1 As String = Data_table_civil_layers.Rows(i).Item("NUME_LAYER")
                            Dim Descr As String = Data_table_civil_layers.Rows(i).Item("DESCRIERE_LAYER")
                            Dim Culoare As String = Data_table_civil_layers.Rows(i).Item("CULOARE")
                            Dim LineW As LineWeight = Data_table_civil_layers.Rows(i).Item("LINEWEIGHT")
                            Dim Plot1 As Boolean = Data_table_civil_layers.Rows(i).Item("PLOT")
                            Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                        Next
                    End If

                    If Data_table_electrical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                        For i = 0 To Data_table_electrical_layers.Rows.Count - 1
                            Dim Ltype As String = Data_table_electrical_layers.Rows(i).Item("LINETYPE")
                            Dim Nume1 As String = Data_table_electrical_layers.Rows(i).Item("NUME_LAYER")
                            Dim Descr As String = Data_table_electrical_layers.Rows(i).Item("DESCRIERE_LAYER")
                            Dim Culoare As String = Data_table_electrical_layers.Rows(i).Item("CULOARE")
                            Dim LineW As LineWeight = Data_table_electrical_layers.Rows(i).Item("LINEWEIGHT")
                            Dim Plot1 As Boolean = Data_table_electrical_layers.Rows(i).Item("PLOT")
                            Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                        Next
                    End If

                    If Data_table_mechanical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                        For i = 0 To Data_table_mechanical_layers.Rows.Count - 1
                            Dim Ltype As String = Data_table_mechanical_layers.Rows(i).Item("LINETYPE")
                            Dim Nume1 As String = Data_table_mechanical_layers.Rows(i).Item("NUME_LAYER")
                            Dim Descr As String = Data_table_mechanical_layers.Rows(i).Item("DESCRIERE_LAYER")
                            Dim Culoare As String = Data_table_mechanical_layers.Rows(i).Item("CULOARE")
                            Dim LineW As LineWeight = Data_table_mechanical_layers.Rows(i).Item("LINEWEIGHT")
                            Dim Plot1 As Boolean = Data_table_mechanical_layers.Rows(i).Item("PLOT")
                            Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                        Next
                    End If

                    If Data_table_pipeline_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                        For i = 0 To Data_table_pipeline_layers.Rows.Count - 1
                            Dim Ltype As String = Data_table_pipeline_layers.Rows(i).Item("LINETYPE")
                            Dim Nume1 As String = Data_table_pipeline_layers.Rows(i).Item("NUME_LAYER")
                            Dim Descr As String = Data_table_pipeline_layers.Rows(i).Item("DESCRIERE_LAYER")
                            Dim Culoare As String = Data_table_pipeline_layers.Rows(i).Item("CULOARE")
                            Dim LineW As LineWeight = Data_table_pipeline_layers.Rows(i).Item("LINEWEIGHT")
                            Dim Plot1 As Boolean = Data_table_pipeline_layers.Rows(i).Item("PLOT")
                            Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                        Next
                    End If

                    If Data_table_mapping_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                        For i = 0 To Data_table_mapping_layers.Rows.Count - 1
                            Dim Ltype As String = Data_table_mapping_layers.Rows(i).Item("LINETYPE")
                            Dim Nume1 As String = Data_table_mapping_layers.Rows(i).Item("NUME_LAYER")
                            Dim Descr As String = Data_table_mapping_layers.Rows(i).Item("DESCRIERE_LAYER")
                            Dim Culoare As String = Data_table_mapping_layers.Rows(i).Item("CULOARE")
                            Dim LineW As LineWeight = Data_table_mapping_layers.Rows(i).Item("LINEWEIGHT")
                            Dim Plot1 As Boolean = Data_table_mapping_layers.Rows(i).Item("PLOT")
                            Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                        Next
                    End If

                    If Data_table_extra_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                        For i = 0 To Data_table_extra_layers.Rows.Count - 1
                            Dim Ltype As String = Data_table_extra_layers.Rows(i).Item("LINETYPE")
                            Dim Nume1 As String = Data_table_extra_layers.Rows(i).Item("NUME_LAYER")
                            Dim Descr As String = Data_table_extra_layers.Rows(i).Item("DESCRIERE_LAYER")
                            Dim Culoare As String = Data_table_extra_layers.Rows(i).Item("CULOARE")
                            Dim LineW As LineWeight = Data_table_extra_layers.Rows(i).Item("LINEWEIGHT")
                            Dim Plot1 As Boolean = Data_table_extra_layers.Rows(i).Item("PLOT")
                            Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                        Next
                    End If
                End If
                afiseaza_butoanele()
            End If

            SET_FILEDIA_TO_1()
        Catch ex As Exception
            Application.SetSystemVariable("FILEDIA", 1)
            afiseaza_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_Load_all_layers_Click(sender As Object, e As EventArgs) Handles Button_Load_all_layers.Click
        Try
            Incarca_toate_linetypes()



            ascunde_butoanele()
            For i = 0 To Data_table_general_layers.Rows.Count - 1
                Dim Ltype As String = Data_table_general_layers.Rows(i).Item("LINETYPE")
                Dim Nume1 As String = Data_table_general_layers.Rows(i).Item("NUME_LAYER")
                Dim Descr As String = Data_table_general_layers.Rows(i).Item("DESCRIERE_LAYER")
                Dim Culoare As String = Data_table_general_layers.Rows(i).Item("CULOARE")
                Dim LineW As LineWeight = Data_table_general_layers.Rows(i).Item("LINEWEIGHT")
                Dim Plot1 As Boolean = Data_table_general_layers.Rows(i).Item("PLOT")
                Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
            Next



            For i = 0 To Data_table_civil_layers.Rows.Count - 1
                Dim Ltype As String = Data_table_civil_layers.Rows(i).Item("LINETYPE")
                Dim Nume1 As String = Data_table_civil_layers.Rows(i).Item("NUME_LAYER")
                Dim Descr As String = Data_table_civil_layers.Rows(i).Item("DESCRIERE_LAYER")
                Dim Culoare As String = Data_table_civil_layers.Rows(i).Item("CULOARE")
                Dim LineW As LineWeight = Data_table_civil_layers.Rows(i).Item("LINEWEIGHT")
                Dim Plot1 As Boolean = Data_table_civil_layers.Rows(i).Item("PLOT")
                Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
            Next



            For i = 0 To Data_table_electrical_layers.Rows.Count - 1
                Dim Ltype As String = Data_table_electrical_layers.Rows(i).Item("LINETYPE")
                Dim Nume1 As String = Data_table_electrical_layers.Rows(i).Item("NUME_LAYER")
                Dim Descr As String = Data_table_electrical_layers.Rows(i).Item("DESCRIERE_LAYER")
                Dim Culoare As String = Data_table_electrical_layers.Rows(i).Item("CULOARE")
                Dim LineW As LineWeight = Data_table_electrical_layers.Rows(i).Item("LINEWEIGHT")
                Dim Plot1 As Boolean = Data_table_electrical_layers.Rows(i).Item("PLOT")
                Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
            Next



            For i = 0 To Data_table_mechanical_layers.Rows.Count - 1
                Dim Ltype As String = Data_table_mechanical_layers.Rows(i).Item("LINETYPE")

                Dim Nume1 As String = Data_table_mechanical_layers.Rows(i).Item("NUME_LAYER")
                Dim Descr As String = Data_table_mechanical_layers.Rows(i).Item("DESCRIERE_LAYER")
                Dim Culoare As String = Data_table_mechanical_layers.Rows(i).Item("CULOARE")
                Dim LineW As LineWeight = Data_table_mechanical_layers.Rows(i).Item("LINEWEIGHT")
                Dim Plot1 As Boolean = Data_table_mechanical_layers.Rows(i).Item("PLOT")
                Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
            Next



            For i = 0 To Data_table_pipeline_layers.Rows.Count - 1
                Dim Ltype As String = Data_table_pipeline_layers.Rows(i).Item("LINETYPE")

                Dim Nume1 As String = Data_table_pipeline_layers.Rows(i).Item("NUME_LAYER")
                Dim Descr As String = Data_table_pipeline_layers.Rows(i).Item("DESCRIERE_LAYER")
                Dim Culoare As String = Data_table_pipeline_layers.Rows(i).Item("CULOARE")
                Dim LineW As LineWeight = Data_table_pipeline_layers.Rows(i).Item("LINEWEIGHT")
                Dim Plot1 As Boolean = Data_table_pipeline_layers.Rows(i).Item("PLOT")
                Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
            Next


            For i = 0 To Data_table_mapping_layers.Rows.Count - 1
                Dim Ltype As String = Data_table_mapping_layers.Rows(i).Item("LINETYPE")
                Dim Nume1 As String = Data_table_mapping_layers.Rows(i).Item("NUME_LAYER")
                Dim Descr As String = Data_table_mapping_layers.Rows(i).Item("DESCRIERE_LAYER")
                Dim Culoare As String = Data_table_mapping_layers.Rows(i).Item("CULOARE")
                Dim LineW As LineWeight = Data_table_mapping_layers.Rows(i).Item("LINEWEIGHT")
                Dim Plot1 As Boolean = Data_table_mapping_layers.Rows(i).Item("PLOT")
                Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
            Next



            For i = 0 To Data_table_extra_layers.Rows.Count - 1
                Dim Ltype As String = Data_table_extra_layers.Rows(i).Item("LINETYPE")
                Dim Nume1 As String = Data_table_extra_layers.Rows(i).Item("NUME_LAYER")
                Dim Descr As String = Data_table_extra_layers.Rows(i).Item("DESCRIERE_LAYER")
                Dim Culoare As String = Data_table_extra_layers.Rows(i).Item("CULOARE")
                Dim LineW As LineWeight = Data_table_extra_layers.Rows(i).Item("LINEWEIGHT")
                Dim Plot1 As Boolean = Data_table_extra_layers.Rows(i).Item("PLOT")
                Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
            Next
            afiseaza_butoanele()




            SET_FILEDIA_TO_1()
        Catch ex As Exception
            Application.SetSystemVariable("FILEDIA", 1)
            afiseaza_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_LOAD_1_LAYER_Click(sender As Object, e As EventArgs) Handles Button_LOAD_1_LAYER.Click
        Try
            Dim Curent_index As Integer = ListBox_LAYERING_name.SelectedIndex
            Incarca_toate_linetypes()

            If Not Curent_index = -1 Then
                ascunde_butoanele()
                If ListBox_LAYERING_name.Items.Count > 0 Then
                    If Data_table_general_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then

                        Dim Ltype As String = Data_table_general_layers.Rows(Curent_index).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_general_layers.Rows(Curent_index).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_general_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_general_layers.Rows(Curent_index).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_general_layers.Rows(Curent_index).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_general_layers.Rows(Curent_index).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)

                    End If

                    If Data_table_civil_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then

                        Dim Ltype As String = Data_table_civil_layers.Rows(Curent_index).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_civil_layers.Rows(Curent_index).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_civil_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_civil_layers.Rows(Curent_index).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_civil_layers.Rows(Curent_index).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_civil_layers.Rows(Curent_index).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)

                    End If

                    If Data_table_electrical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then

                        Dim Ltype As String = Data_table_electrical_layers.Rows(Curent_index).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_electrical_layers.Rows(Curent_index).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_electrical_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_electrical_layers.Rows(Curent_index).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_electrical_layers.Rows(Curent_index).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_electrical_layers.Rows(Curent_index).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)

                    End If

                    If Data_table_mechanical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then

                        Dim Ltype As String = Data_table_mechanical_layers.Rows(Curent_index).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_mechanical_layers.Rows(Curent_index).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_mechanical_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_mechanical_layers.Rows(Curent_index).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_mechanical_layers.Rows(Curent_index).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_mechanical_layers.Rows(Curent_index).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)

                    End If

                    If Data_table_pipeline_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then

                        Dim Ltype As String = Data_table_pipeline_layers.Rows(Curent_index).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_pipeline_layers.Rows(Curent_index).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_pipeline_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_pipeline_layers.Rows(Curent_index).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_pipeline_layers.Rows(Curent_index).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_pipeline_layers.Rows(Curent_index).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)

                    End If

                    If Data_table_mapping_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then

                        Dim Ltype As String = Data_table_mapping_layers.Rows(Curent_index).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_mapping_layers.Rows(Curent_index).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_mapping_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_mapping_layers.Rows(Curent_index).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_mapping_layers.Rows(Curent_index).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_mapping_layers.Rows(Curent_index).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)

                    End If

                    If Data_table_extra_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then

                        Dim Ltype As String = Data_table_extra_layers.Rows(Curent_index).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_extra_layers.Rows(Curent_index).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_extra_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_extra_layers.Rows(Curent_index).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_extra_layers.Rows(Curent_index).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_extra_layers.Rows(Curent_index).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)

                    End If
                End If
                afiseaza_butoanele()
            End If


            SET_FILEDIA_TO_1()

        Catch ex As Exception
            Application.SetSystemVariable("FILEDIA", 1)
            afiseaza_butoanele()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_general_LAYERING_Click(sender As Object, e As EventArgs) Handles Button_GENERAL_LAYERING.Click
        Try
            ListBox_LAYERING_name.Items.Clear()
            For i = 0 To Data_table_general_layers.Rows.Count - 1
                ListBox_LAYERING_name.Items.Add(Data_table_general_layers.Rows(i).Item("NUME_LAYER"))
            Next
            ListBox_LAYERING_name.Visible = True
            TextBox_description.Text = ""
            TextBox_description.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = True
            Button_LOAD_1_LAYER.Visible = True
            CheckBox_my_list.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_CIVIL_LAYERYING_Click(sender As Object, e As EventArgs) Handles Button_CIVIL_LAYERYING.Click
        Try
            ListBox_LAYERING_name.Items.Clear()
            For i = 0 To Data_table_civil_layers.Rows.Count - 1
                ListBox_LAYERING_name.Items.Add(Data_table_civil_layers.Rows(i).Item("NUME_LAYER"))
            Next
            ListBox_LAYERING_name.Visible = True
            TextBox_description.Text = ""
            TextBox_description.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = True
            Button_LOAD_1_LAYER.Visible = True
            CheckBox_my_list.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_ELECTRICAL_LAYERING_Click(sender As Object, e As EventArgs) Handles Button_ELECTRICAL_LAYERING.Click
        Try
            ListBox_LAYERING_name.Items.Clear()
            For i = 0 To Data_table_electrical_layers.Rows.Count - 1
                ListBox_LAYERING_name.Items.Add(Data_table_electrical_layers.Rows(i).Item("NUME_LAYER"))
            Next
            ListBox_LAYERING_name.Visible = True
            TextBox_description.Text = ""
            TextBox_description.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = True
            Button_LOAD_1_LAYER.Visible = True
            CheckBox_my_list.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_MECHANICAL_LAYERING_Click(sender As Object, e As EventArgs) Handles Button_MECHANICAL_LAYERING.Click
        Try
            ListBox_LAYERING_name.Items.Clear()
            For i = 0 To Data_table_mechanical_layers.Rows.Count - 1
                ListBox_LAYERING_name.Items.Add(Data_table_mechanical_layers.Rows(i).Item("NUME_LAYER"))
            Next
            ListBox_LAYERING_name.Visible = True
            TextBox_description.Text = ""
            TextBox_description.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = True
            Button_LOAD_1_LAYER.Visible = True
            CheckBox_my_list.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_PIPELINE_LAYERING_Click(sender As Object, e As EventArgs) Handles Button_PIPELINE_LAYERING.Click
        Try
            ListBox_LAYERING_name.Items.Clear()
            For i = 0 To Data_table_pipeline_layers.Rows.Count - 1
                ListBox_LAYERING_name.Items.Add(Data_table_pipeline_layers.Rows(i).Item("NUME_LAYER"))
            Next
            ListBox_LAYERING_name.Visible = True
            TextBox_description.Text = ""
            TextBox_description.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = True
            Button_LOAD_1_LAYER.Visible = True
            CheckBox_my_list.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_MAPPING_LAYERING_Click(sender As Object, e As EventArgs) Handles Button_MAPPING_LAYERING.Click
        Try
            ListBox_LAYERING_name.Items.Clear()
            For i = 0 To Data_table_mapping_layers.Rows.Count - 1
                ListBox_LAYERING_name.Items.Add(Data_table_mapping_layers.Rows(i).Item("NUME_LAYER"))
            Next
            ListBox_LAYERING_name.Visible = True
            TextBox_description.Text = ""
            TextBox_description.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = True
            Button_LOAD_1_LAYER.Visible = True
            CheckBox_my_list.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button_EXTRA_LAYERING_Click(sender As Object, e As EventArgs) Handles Button_EXTRA_LAYERING.Click
        Try
            ListBox_LAYERING_name.Items.Clear()
            For i = 0 To Data_table_extra_layers.Rows.Count - 1
                ListBox_LAYERING_name.Items.Add(Data_table_extra_layers.Rows(i).Item("NUME_LAYER"))
            Next
            ListBox_LAYERING_name.Visible = True
            TextBox_description.Text = ""
            TextBox_description.Visible = False
            Button_LOAD_GRUP_LAYERS.Visible = True
            Button_LOAD_1_LAYER.Visible = True
            CheckBox_my_list.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub CheckBox_my_list_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_my_list.CheckedChanged
        Try
            If CheckBox_my_list.Checked = True Then
                ListBox_my_list.Visible = True
                Me.Width = 948
                Button_load_from_my_list.Visible = True
                Button_CREATE_UPDATE_MY_LIST.Visible = True
                Button_CREATE_UPDATE_MY_LIST.Visible = True
                Button_Read_my_list.Visible = True
            Else
                ListBox_my_list.Visible = False
                Button_load_from_my_list.Visible = False
                Button_CREATE_UPDATE_MY_LIST.Visible = False
                Button_CREATE_UPDATE_MY_LIST.Visible = False
                Button_Read_my_list.Visible = False
                Me.Width = 752
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub




    Private Sub Button_CREATE_UPDATE_MY_LIST_Click(sender As Object, e As EventArgs) Handles Button_CREATE_UPDATE_MY_LIST.Click
        ascunde_butoanele()
        Try
            If Data_table_my_list_layers.Rows.Count > 0 Then
                If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(Locatie1) = False Then
                    Microsoft.VisualBasic.FileIO.FileSystem.CreateDirectory(Locatie1)
                End If
                Dim Fisier As String = Locatie1 & "\My_list.txt"
                If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Fisier) = False Then
                    Dim fs As IO.FileStream = System.IO.File.Create(Fisier)
                    fs.Close()
                End If
                Dim StreamWriter1 As New System.IO.StreamWriter(Fisier)
                Using StreamWriter1
                    For i = 0 To Data_table_my_list_layers.Rows.Count - 1
                        If Not i = Data_table_my_list_layers.Rows.Count - 1 Then
                            StreamWriter1.Write(Data_table_my_list_layers.Rows(i).Item("NUME_LAYER") & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("DESCRIERE_LAYER") & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("CULOARE") & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("LINEWEIGHT").ToString & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("PLOT").ToString & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("LINETYPE") & vbCrLf)
                        Else
                            StreamWriter1.Write(Data_table_my_list_layers.Rows(i).Item("NUME_LAYER") & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("DESCRIERE_LAYER") & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("CULOARE") & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("LINEWEIGHT").ToString & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("PLOT").ToString & "~" & _
                                                Data_table_my_list_layers.Rows(i).Item("LINETYPE"))
                        End If
                    Next
                    StreamWriter1.Close()
                End Using



            End If



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele()
    End Sub

    Private Sub ListBox_my_list_Click(sender As Object, e As EventArgs) Handles ListBox_my_list.Click
        ascunde_butoanele()
        Try
            Dim Curent_index As Integer = ListBox_my_list.SelectedIndex
            If Not Curent_index = -1 Then
                ListBox_my_list.Items.RemoveAt(Curent_index)
                If Data_table_my_list_layers.Rows.Count >= Curent_index Then
                    Data_table_my_list_layers.Rows(Curent_index).Delete()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele()
    End Sub


    Private Sub Button_Read_my_list_Click(sender As Object, e As EventArgs) Handles Button_Read_my_list.Click
        ascunde_butoanele()
        Try
            ListBox_my_list.Items.Clear()
            
            Dim Fisier As String = Locatie1 & "\My_list.txt"
            If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Fisier) = True Then
                Data_table_my_list_layers.Rows.Clear()
                Using Reader1 As New System.IO.StreamReader(Fisier)
                    Dim Line1 As String
                    While Reader1.Peek > 0
                        Line1 = Reader1.ReadLine
                        Dim Cuvinte() As String = Line1.Split("~")
                        Data_table_my_list_layers.Rows.Add()
                        Dim Index_my_list As Integer = Data_table_my_list_layers.Rows.Count - 1
                        Data_table_my_list_layers.Rows(Index_my_list).Item("NUME_LAYER") = Cuvinte(0)
                        ListBox_my_list.Items.Add(Cuvinte(0))
                        Data_table_my_list_layers.Rows(Index_my_list).Item("DESCRIERE_LAYER") = Cuvinte(1)
                        Data_table_my_list_layers.Rows(Index_my_list).Item("CULOARE") = CInt(Cuvinte(2))
                        Data_table_my_list_layers.Rows(Index_my_list).Item("LINETYPE") = Cuvinte(5)
                        Select Case Cuvinte(4).ToUpper
                            Case "TRUE"
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = True
                            Case "FALSE"
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = False
                            Case Else
                                Data_table_my_list_layers.Rows(Index_my_list).Item("PLOT") = True
                        End Select
                        Select Case Cuvinte(3)
                            Case 0
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight000
                            Case 5
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight005
                            Case 13
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight013
                            Case 15
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight015
                            Case 18
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight018
                            Case 20
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight020
                            Case 25
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight025
                            Case 30
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight030
                            Case 35
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight035
                            Case 40
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight040
                            Case 50
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight050
                            Case 53
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight053
                            Case 60
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight060
                            Case 70
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight070
                            Case 80
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight080
                            Case 90
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight090
                            Case 100
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight100
                            Case 106
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight106
                            Case 120
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight120
                            Case 158
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight158
                            Case 200
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight200
                            Case 201
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.LineWeight211
                            Case Else
                                Data_table_my_list_layers.Rows(Index_my_list).Item("LINEWEIGHT") = LineWeight.ByLineWeightDefault
                        End Select


                    End While


                End Using
            Else
                MsgBox("You do not have a predifined list on your PC")
            End If



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        afiseaza_butoanele()
    End Sub

    Private Sub Button_load_from_my_list_Click(sender As Object, e As EventArgs) Handles Button_load_from_my_list.Click
        ascunde_butoanele()
        Try
            If Data_table_my_list_layers.Rows.Count > 0 Then
                Creaza_Dim_style_AND_text_style_ROMANS()
                For i = 0 To Data_table_my_list_layers.Rows.Count - 1
                    If IsDBNull(Data_table_my_list_layers.Rows(i).Item("NUME_LAYER")) = False And IsDBNull(Data_table_my_list_layers.Rows(i).Item("DESCRIERE_LAYER")) = False _
                        And IsDBNull(Data_table_my_list_layers.Rows(i).Item("CULOARE")) = False And IsDBNull(Data_table_my_list_layers.Rows(i).Item("LINEWEIGHT")) = False _
                        And IsDBNull(Data_table_my_list_layers.Rows(i).Item("PLOT")) = False And IsDBNull(Data_table_my_list_layers.Rows(i).Item("LINETYPE")) = False Then
                        Dim Ltype As String = Data_table_my_list_layers.Rows(i).Item("LINETYPE")
                        Dim Nume1 As String = Data_table_my_list_layers.Rows(i).Item("NUME_LAYER")
                        Dim Descr As String = Data_table_my_list_layers.Rows(i).Item("DESCRIERE_LAYER")
                        Dim Culoare As String = Data_table_my_list_layers.Rows(i).Item("CULOARE")
                        Dim LineW As LineWeight = Data_table_my_list_layers.Rows(i).Item("LINEWEIGHT")
                        Dim Plot1 As Boolean = Data_table_my_list_layers.Rows(i).Item("PLOT")
                        Creaza_layer_cu_linetype_si_lineweight(Nume1, Culoare, Ltype, LineW, Descr, Plot1, True)
                    Else
                        MsgBox("The list is not properly created, Define it again")
                    End If
                Next
            Else
                MsgBox("First press read my list and then press load my layers")
            End If
            SET_FILEDIA_TO_1()
        Catch ex As Exception
            MsgBox(ex.Message)
            SET_FILEDIA_TO_1()
        End Try
        afiseaza_butoanele()
    End Sub

    Private Sub Button_fix_xrefs_ltypes_Click(sender As Object, e As EventArgs) Handles Button_fix_xrefs_ltypes.Click
        ascunde_butoanele()
        Creaza_Dim_style_AND_text_style_ROMANS()
        Try
            Dim ThisDrawing As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using lock As DocumentLock = ThisDrawing.LockDocument
                Using Trans1 As Transaction = ThisDrawing.Database.TransactionManager.StartTransaction
                    Dim LayerTable1 As Autodesk.AutoCAD.DatabaseServices.LayerTable = Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    Dim LineTable1 As Autodesk.AutoCAD.DatabaseServices.LinetypeTable = Trans1.GetObject(ThisDrawing.Database.LinetypeTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                    Dim Colectie_nume_linetypes As New Specialized.StringCollection
                    For Each LtypeId As ObjectId In LineTable1
                        Dim Linetype1 As LinetypeTableRecord = LtypeId.GetObject(OpenMode.ForRead)
                        If Linetype1.Name.Contains("|") = False And Linetype1.Name.Contains("$") = False Then
                            Colectie_nume_linetypes.Add(Linetype1.Name)
                        End If
                    Next

                    If Colectie_nume_linetypes.Count > 0 Then
                        For Each LayerId As ObjectId In LayerTable1
                            Dim Layer1 As LayerTableRecord = Trans1.GetObject(LayerId, OpenMode.ForRead)
                            Dim Linetype1 As LinetypeTableRecord = Layer1.LinetypeObjectId.GetObject(OpenMode.ForRead)
                            Dim Nume_linie_layer As String = Linetype1.Name
                            If Nume_linie_layer.Contains("|") = True Then
                                Dim Pozitie1 As Integer = InStr(Nume_linie_layer, "|")
                                Dim Nume_linie_Xref As String = Strings.Right(Nume_linie_layer, Len(Nume_linie_layer) - Pozitie1)
                                If Colectie_nume_linetypes.Contains(Nume_linie_Xref) = True Then
                                    Layer1.UpgradeOpen()
                                    Layer1.LinetypeObjectId = LineTable1.Item(Nume_linie_Xref)
                                End If
                            End If
                        Next
                        Trans1.Commit()
                    End If
                End Using
            End Using
            SET_FILEDIA_TO_1()
            afiseaza_butoanele()
        Catch ex As Exception
            afiseaza_butoanele()
            Application.SetSystemVariable("FILEDIA", 1)
            MsgBox(ex.Message)
        End Try
        SET_FILEDIA_TO_1()
    End Sub



    Private Sub ListBox_LAYERING_name_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox_LAYERING_name.SelectedIndexChanged
        Try
            Dim Curent_index As Integer = ListBox_LAYERING_name.SelectedIndex
            If Not Curent_index = -1 Then
                TextBox_description.Visible = True

                If Data_table_general_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                    Dim linew As Double = CDbl(Data_table_general_layers.Rows(Curent_index).Item("LINEWEIGHT").ToString) / 100

                    TextBox_description.Text = "COLOR = " & Data_table_general_layers.Rows(Curent_index).Item("CULOARE") & vbCrLf _
                                            & "Linetype = " & Data_table_general_layers.Rows(Curent_index).Item("LINETYPE") & vbCrLf _
                                            & "Lineweight = " & linew & vbCrLf & vbCrLf & Data_table_general_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                End If

                If Data_table_civil_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                    Dim linew As Double = CDbl(Data_table_civil_layers.Rows(Curent_index).Item("LINEWEIGHT").ToString) / 100
                    TextBox_description.Text = "COLOR = " & Data_table_civil_layers.Rows(Curent_index).Item("CULOARE") & vbCrLf _
                                            & "Linetype = " & Data_table_civil_layers.Rows(Curent_index).Item("LINETYPE") & vbCrLf _
                                            & "Lineweight = " & linew & vbCrLf & vbCrLf & Data_table_civil_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                End If

                If Data_table_electrical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                    Dim linew As Double = CDbl(Data_table_electrical_layers.Rows(Curent_index).Item("LINEWEIGHT").ToString) / 100
                    TextBox_description.Text = "COLOR = " & Data_table_electrical_layers.Rows(Curent_index).Item("CULOARE") & vbCrLf _
                                            & "Linetype = " & Data_table_electrical_layers.Rows(Curent_index).Item("LINETYPE") & vbCrLf _
                                            & "Lineweight = " & linew & vbCrLf & vbCrLf & Data_table_electrical_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                End If

                If Data_table_mechanical_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                    Dim linew As Double = CDbl(Data_table_mechanical_layers.Rows(Curent_index).Item("LINEWEIGHT").ToString) / 100
                    TextBox_description.Text = "COLOR = " & Data_table_mechanical_layers.Rows(Curent_index).Item("CULOARE") & vbCrLf _
                                            & "Linetype = " & Data_table_mechanical_layers.Rows(Curent_index).Item("LINETYPE") & vbCrLf _
                                            & "Lineweight = " & linew & vbCrLf & vbCrLf & Data_table_mechanical_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                End If

                If Data_table_pipeline_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                    Dim linew As Double = CDbl(Data_table_pipeline_layers.Rows(Curent_index).Item("LINEWEIGHT").ToString) / 100
                    TextBox_description.Text = "COLOR = " & Data_table_pipeline_layers.Rows(Curent_index).Item("CULOARE") & vbCrLf _
                                            & "Linetype = " & Data_table_pipeline_layers.Rows(Curent_index).Item("LINETYPE") & vbCrLf _
                                            & "Lineweight = " & linew & vbCrLf & vbCrLf & Data_table_pipeline_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                End If

                If Data_table_mapping_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                    Dim linew As Double = CDbl(Data_table_mapping_layers.Rows(Curent_index).Item("LINEWEIGHT").ToString) / 100
                    TextBox_description.Text = "COLOR = " & Data_table_mapping_layers.Rows(Curent_index).Item("CULOARE") & vbCrLf _
                                            & "Linetype = " & Data_table_mapping_layers.Rows(Curent_index).Item("LINETYPE") & vbCrLf _
                                            & "Lineweight = " & linew & vbCrLf & vbCrLf & Data_table_mapping_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                End If

                If Data_table_extra_layers.Rows(0).Item("NUME_LAYER") = ListBox_LAYERING_name.Items(0) Then
                    Dim linew As Double = CDbl(Data_table_extra_layers.Rows(Curent_index).Item("LINEWEIGHT").ToString) / 100
                    TextBox_description.Text = "COLOR = " & Data_table_extra_layers.Rows(Curent_index).Item("CULOARE") & vbCrLf _
                                            & "Linetype = " & Data_table_extra_layers.Rows(Curent_index).Item("LINETYPE") & vbCrLf _
                                            & "Lineweight = " & linew & vbCrLf & vbCrLf & Data_table_extra_layers.Rows(Curent_index).Item("DESCRIERE_LAYER")
                End If

            End If




        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class