Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime

Module ExtensionMethods

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SynchronizeAttributes(target As BlockTableRecord)
        If target Is Nothing Then
            Throw New ArgumentNullException("btr")
        End If

        Dim tr As Transaction = target.Database.TransactionManager.TopTransaction
        If tr Is Nothing Then
            Throw New Exception(ErrorStatus.NoActiveTransactions)
        End If

        Dim attDefClass As RXClass = RXClass.GetClass(GetType(AttributeDefinition))
        Dim attDefs As New List(Of AttributeDefinition)()
        For Each id As ObjectId In target
            If id.ObjectClass = attDefClass Then
                Dim attDef As AttributeDefinition = DirectCast(tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeDefinition)
                attDefs.Add(attDef)
            End If
        Next
        For Each id As ObjectId In target.GetBlockReferenceIds(True, False)
            Dim br As BlockReference = DirectCast(tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite), BlockReference)
            br.ResetAttributes(attDefs)
        Next
        If target.IsDynamicBlock Then
            For Each id As ObjectId In target.GetAnonymousBlockIds()
                Dim btr As BlockTableRecord = DirectCast(tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTableRecord)
                For Each brId As ObjectId In btr.GetBlockReferenceIds(True, False)
                    Dim br As BlockReference = DirectCast(tr.GetObject(brId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite), BlockReference)
                    br.ResetAttributes(attDefs)
                Next
            Next
        End If
    End Sub



    <System.Runtime.CompilerServices.Extension()> _
    Private Sub ResetAttributes(br As BlockReference, attDefs As List(Of AttributeDefinition))
        Dim tm As TransactionManager = br.Database.TransactionManager
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
            Dim br As BlockReference = DirectCast(Tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite), BlockReference)
            br.ResetAttributes(attDefs)
        Next


        If target.IsDynamicBlock Then
            For Each id As ObjectId In target.GetAnonymousBlockIds()
                Dim btr1 As BlockTableRecord = DirectCast(Tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTableRecord)
                For Each brId As ObjectId In btr1.GetBlockReferenceIds(True, False)
                    Dim br As BlockReference = DirectCast(Tr.GetObject(brId, Autodesk.AutoCAD.DatabaseServices.OpenMode.Forwrite), BlockReference)
                    br.ResetAttributes_db_diferit(attDefs)
                Next
            Next
        End If
    End Sub


    <System.Runtime.CompilerServices.Extension()> _
    Private Sub ResetAttributes_db_diferit(br As BlockReference, attDefs As List(Of AttributeDefinition))
        Dim tm As TransactionManager = br.Database.TransactionManager
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


End Module
