Imports ACSMCOMPONENTS20Lib
Imports Autodesk.AutoCAD.Runtime
Imports System.Math
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Sheet_set_manager_commands
    <CommandMethod("CRRR")> _
    Public Sub Open_sheet_set()
        Dim SheetSet_manager As New AcSmSheetSetMgr
        Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase("C:\Users\pop70694\Documents\WORK FILES\2015-05-08 plat generator\SEG_H_314.00 - Standard\KinderMorgan-NED-SegA_LandPlats.dst", False)
        Dim sheetSet As AcSmSheetSet
        sheetSet = SheetSet_database.GetSheetSet()

        ' Attempt to lock the database
        If LockDatabase(SheetSet_database, True) = True Then
            Dim EnumSheets As IAcSmEnumComponent = sheetSet.GetSheetEnumerator()
            Dim smComponent As IAcSmComponent
            Dim sheet1 As IAcSmSheet
            'Dim sheet2 As IAcSmSheet2

            smComponent = EnumSheets.Next()
            While True
                If smComponent Is Nothing Then
                    Exit While
                End If

                sheet1 = TryCast(smComponent, IAcSmSheet)

                If sheet1.GetName.Contains("314") = True Then
                    SetCustomProperty(sheetSet, "BBBB", "BB1", PropertyFlags.CUSTOM_SHEET_PROP)
                Else
                    SetCustomProperty(sheetSet, "BBBB", "BB2", PropertyFlags.CUSTOM_SHEET_PROP)
                End If

                'MsgBox(sheet1.GetName.ToString)
                ' To access the revision number, 
                ' Revision date, Purpose and Category,
                ' cast it as IAcSmSheet2
                'sheet2 = TryCast(smComponent, IAcSmSheet2)
                'If sheet2 IsNot Nothing Then
                'sheet2.SetRevisionNumber("XXX111")
                'End If

                smComponent = EnumSheets.Next()
            End While

            ' Create a sheet set property
            'SetCustomProperty(sheetSet, "AAAA", "AA1", PropertyFlags.CUSTOM_SHEETSET_PROP)
            ' Create sheet properties
            ' SetCustomProperty(sheetSet, "BBBB", "BB1", PropertyFlags.CUSTOM_SHEET_PROP)
            ' Unlock the database
            LockDatabase(SheetSet_database, False)
        Else
            ' Display error message
            MsgBox("Sheet set could not be opened for write.")
        End If

        ' Close the sheet set
        SheetSet_manager.Close(SheetSet_database)


    End Sub

    <CommandMethod("CRRR1")> _
    Public Sub Open_sheet_set1()
        Dim SheetSet_manager As New AcSmSheetSetMgr
        Dim SheetSet_database As AcSmDatabase = SheetSet_manager.OpenDatabase("C:\Users\pop70694\Documents\WORK FILES\2015-05-08 plat generator\SEG_H_314.00 - Standard\KinderMorgan-NED-SegA_LandPlats.dst", False)
        Dim sheetSet As AcSmSheetSet
        sheetSet = SheetSet_database.GetSheetSet()

        ' Attempt to lock the database
        If LockDatabase(SheetSet_database, True) = True Then
            Dim sheetSetFolder As String
            'sheetSetFolder = Mid(SheetSet_database.GetFileName(), 1, InStrRev(SheetSet_database.GetFileName(), "\"))
            sheetSetFolder = "C:\Users\pop70694\Documents\WORK FILES\2015-05-08 plat generator\test\"


            SetSheetSetDefaults(SheetSet_database, _
                    "KinderMorgan-NED-SegA_LandPlats", "KinderMorgan-NED-SegA_LandPlats", sheetSetFolder, "C:\Users\pop70694\Documents\WORK FILES\2015-05-08 plat generator\sheet set platt\2015_05_26_LandPlat_Template\KinderMorgan_NED_Seg_A_PlatTemplate.dwt", "Layout1")
            AddSheet(SheetSet_database, "SEG_A_318.00", "SEG_A_318.00", "SEG_A_318.00", "1 of 1")


            ' Create a sheet set property
            'SetCustomProperty(sheetSet, "AAAA", "AA1", PropertyFlags.CUSTOM_SHEETSET_PROP)
            ' Create sheet properties
            ' SetCustomProperty(sheetSet, "BBBB", "BB1", PropertyFlags.CUSTOM_SHEET_PROP)
            ' Unlock the database
            LockDatabase(SheetSet_database, False)
        Else
            ' Display error message
            MsgBox("Sheet set could not be opened for write.")
        End If

        ' Close the sheet set
        SheetSet_manager.Close(SheetSet_database)


    End Sub

    Private Sub SetCustomProperty(ByVal owner As IAcSmPersist, _
                                      ByVal propertyName As String, _
                                      ByVal propertyValue As Object, _
                                      ByVal sheetSetFlag As PropertyFlags)

        ' Create a reference to the Custom Property Bag

        Dim customPropertyBag As AcSmCustomPropertyBag

        If owner.GetTypeName() = "AcSmSheet" Then
            Dim sheet As AcSmSheet = owner
            customPropertyBag = sheet.GetCustomPropertyBag()
        Else
            Dim sheetSet As AcSmSheetSet = owner
            customPropertyBag = sheetSet.GetCustomPropertyBag()
        End If
        ' Create a reference to a Custom Property Value
        Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
        customPropertyValue.InitNew(owner)
        ' Set the flag for the property
        customPropertyValue.SetFlags(sheetSetFlag)
        ' Set the value for the property
        customPropertyValue.SetValue(propertyValue)
        ' Create the property
        customPropertyBag.SetProperty(propertyName, customPropertyValue)
    End Sub
    ' Used to lock/unlock a sheet set database
    Private Function LockDatabase(ByRef database As AcSmDatabase, _
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
    ' Used to add a sheet to a sheet set or subset
    ' Note: This function is dependent on a Default Template and Storage location
    ' being set for the sheet set or subset.
    Private Function AddSheet(ByVal component As IAcSmComponent, _
                              ByVal name As String, _
                              ByVal description As String, _
                              ByVal title As String, _
                              ByVal number As String) As AcSmSheet

        Dim sheet As AcSmSheet

        ' Check to see if the component is a sheet set or subset, 
        ' and create the new sheet based on the component's type
        If component.GetTypeName = "AcSmSubset" Then
            Dim subset As AcSmSubset = component
            sheet = subset.AddNewSheet(name, description)

            ' Add the sheet as the first one in the subset
            subset.InsertComponent(sheet, Nothing)
        Else
            sheet = component.GetDatabase().GetSheetSet().AddNewSheet(name, _
                                                                      description)

            ' Add the sheet as the first one in the sheet set
            component.GetDatabase().GetSheetSet().InsertComponent(sheet, Nothing)
        End If

        ' Set the number and title of the sheet
        sheet.SetNumber(number)
        sheet.SetTitle(title)

        AddSheet = sheet
    End Function
    Private Function ImportASheet(ByVal component As IAcSmComponent, _
                                  ByVal title As String, _
                                  ByVal description As String, _
                                  ByVal number As String, _
                                  ByVal fileName As String, _
                                  ByVal layout As String) As AcSmSheet

        Dim sheet As AcSmSheet

        ' Create a reference to a Layout Reference object
        Dim layoutReference As New AcSmAcDbLayoutReference
        layoutReference.InitNew(component)

        ' Set the layout and drawing file to use for the sheet
        layoutReference.SetFileName(fileName)
        layoutReference.SetName(layout)

        ' Import the sheet into the sheet set
        ' Check to see if the Component is a Subset or Sheet Set
        If component.GetTypeName = "AcSmSubset" Then
            Dim subset As AcSmSubset = component

            sheet = subset.ImportSheet(layoutReference)
            subset.InsertComponent(sheet, Nothing)
        Else
            Dim sheetSetDatabase As AcSmDatabase = component

            sheet = sheetSetDatabase.GetSheetSet().ImportSheet(layoutReference)
            sheetSetDatabase.GetSheetSet().InsertComponent(sheet, Nothing)
        End If

        ' Set the properties of the sheet
        sheet.SetDesc(description)
        sheet.SetTitle(title)
        sheet.SetNumber(number)

        ImportASheet = sheet
    End Function
    Private Sub SetSheetSetDefaults(ByVal sheetSetDatabase As AcSmDatabase, _
                                    ByVal name As String, _
                                  ByVal description As String, _
                                   ByVal newSheetLocation As String, _
                                  Optional ByVal newSheetDWTLocation As String = "", _
                                  Optional ByVal newSheetDWTLayout As String = "", _
                                  Optional ByVal promptForDWT As Boolean = False)

        ' Set the Name and Description for the sheet set
        sheetSetDatabase.GetSheetSet().SetName(name)
        sheetSetDatabase.GetSheetSet().SetDesc(description)

        ' Check to see if a Storage Location was provided
        If newSheetLocation <> "" Then
            ' Get the folder the sheet set is stored in
            Dim sheetSetFolder As String
            sheetSetFolder = Mid(sheetSetDatabase.GetFileName(), 1, InStrRev(sheetSetDatabase.GetFileName(), "\"))

            ' Create a reference to a File Reference object
            Dim fileReference As IAcSmFileReference
            fileReference = sheetSetDatabase.GetSheetSet().GetNewSheetLocation()

            ' Set the default storage location based on the location of the sheet set
            fileReference.SetFileName(newSheetLocation)

            ' Set the new Sheet location for the sheet set
            sheetSetDatabase.GetSheetSet().SetNewSheetLocation(fileReference)
        End If

        ' Check to see if a Template was provided
        If newSheetDWTLocation <> "" Then
            ' Set the Default Template for the sheet set
            Dim layoutReference As AcSmAcDbLayoutReference
            layoutReference = sheetSetDatabase.GetSheetSet().GetDefDwtLayout()

            ' Set the template location and name of the layout 
            ' for the Layout Reference object
            layoutReference.SetFileName(newSheetDWTLocation)
            layoutReference.SetName(newSheetDWTLayout)

            ' Set the Layout Reference for the sheet set
            sheetSetDatabase.GetSheetSet().SetDefDwtLayout(layoutReference)
        End If

        ' Set the Prompt for Template option of the subset
        sheetSetDatabase.GetSheetSet().SetPromptForDwt(promptForDWT)
    End Sub

End Class

