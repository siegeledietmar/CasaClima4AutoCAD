' (C) Copyright 2011 by  
'
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Windows

Imports Autodesk.AutoCAD.Windows

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(casaclima.MyCommands))>
Namespace casaclima

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class MyCommands

        ' The CommandMethod attribute can be applied to any public  member 
        ' function of any public class.
        ' The function should take no arguments and return nothing.
        ' If the method is an instance member then the enclosing class is 
        ' instantiated for each document. If the member is a static member then
        ' the enclosing class is NOT instantiated.
        '
        ' NOTE: CommandMethod has overloads where you can provide helpid and
        ' context menu.

        ' Modal Command with localized name
        ' AutoCAD will search for a resource string with Id "MyCommandLocal" in the 
        ' same namespace as this command class. 
        ' If a resource string is not found, then the string "MyLocalCommand" is used 
        ' as the localized command name.
        ' To view/edit the resx file defining the resource strings for this command, 
        ' * click the 'Show All Files' button in the Solution Explorer;
        ' * expand the tree node for myCommands.vb;
        ' * and double click on myCommands.resx

        ' Modal Command with pickfirst selection
        ' <CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal + CommandFlags.UsePickSet)>
        <CommandMethod("CCarea", CommandFlags.Modal + CommandFlags.UsePickSet)>
        Public Sub CCarea() ' This method can have any name
            '' Get the editor, current document and database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                '' Request for objects to be selected in the drawing area
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()

                '' If the prompt status is OK, objects were selected
                If acSSPrompt.Status = PromptStatus.OK Then
                    'open the current space (can be any BTR, usually model space)
                    Dim mSpace As BlockTableRecord = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)

                    Dim acSSet As SelectionSet = acSSPrompt.Value

                    '' Step through the objects in the selection set
                    For Each acSSObj As SelectedObject In acSSet
                        '' Check to make sure a valid SelectedObject object was returned
                        If Not IsDBNull(acSSObj) Then
                            '' Open the selected object for write
                            Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)

                            If Not IsDBNull(acEnt) Then
                                Dim aclayer As String = acEnt.Layer

                                'ed.WriteMessage(acEnt.GetType.ToString())

                                Dim polyline As Polyline = TryCast(acEnt, Polyline)
                                If (polyline <> Nothing) Then
                                    'Dim pointOptions As New PromptPointOptions("Punkt für Beschriftung: ")
                                    'Dim pointResult As PromptPointResult = ed.GetPoint(pointOptions)
                                    'Dim selectedPoint As Point3d = pointResult.Value

                                    Dim extents As Extents3d = polyline.GeometricExtents
                                    Dim selectedPoint As Point3d = extents.MinPoint + (extents.MaxPoint - extents.MinPoint) / 2.0

                                    Dim area As Double = polyline.Area

                                    '' Adds all to an object id array
                                    Dim acObjIdColl As ObjectIdCollection = New ObjectIdCollection()
                                    acObjIdColl.Add(polyline.ObjectId)

                                    Using objText As MText = New MText
                                        '' Specify the insertion point of the MText object
                                        objText.Location = selectedPoint

                                        '' Set the text string for the MText object
                                        objText.Contents = Math.Round(area, 2).ToString & " qm"

                                        '' Set the text style for the MText object
                                        objText.TextStyleId = acCurDb.Textstyle
                                        objText.TextHeight = 0.25
                                        objText.Layer = aclayer
                                        Try
                                            objText.Attachment = AttachmentPoint.MiddleCenter
                                        Catch ex As Exception

                                        End Try

                                        '' Appends the new MText object to model space
                                        mSpace.AppendEntity(objText)

                                        '' Appends to new MText object to the active transaction
                                        acTrans.AddNewlyCreatedDBObject(objText, True)

                                        Dim acObjIdCol2 As ObjectIdCollection = New ObjectIdCollection()
                                        acObjIdCol2.Add(objText.ObjectId)

                                        Using objHatch As Hatch = New Hatch
                                            mSpace.AppendEntity(objHatch)
                                            objHatch.PatternScale = 0.1
                                            objHatch.Layer = aclayer
                                            acTrans.AddNewlyCreatedDBObject(objHatch, True)

                                            '' Set the properties of the hatch object
                                            '' Associative must be set after the hatch object is appended to the 
                                            '' block table record and before AppendLoop
                                            objHatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31")
                                            objHatch.Associative = True
                                            objHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl)
                                            objHatch.AppendLoop(HatchLoopTypes.TextIsland, acObjIdCol2)
                                            objHatch.EvaluateHatch(True)
                                        End Using
                                    End Using
                                Else
                                    ed.WriteMessage("\n Selected object not a polyline!")
                                End If


                                'ed.WriteMessage("Hello World")
                            End If
                        End If
                    Next


                    ''Create and configure a new line
                    'Dim newLine As New Line
                    'newLine.StartPoint = selectedPoint
                    'newLine.EndPoint = New Point3d(selectedPoint.Coordinate(0) + 10, selectedPoint.Coordinate(1) + 10, selectedPoint.Coordinate(2))

                    ''Append to model space
                    'mSpace.AppendEntity(newLine)

                    ''Inform the transaction
                    'acTrans.AddNewlyCreatedDBObject(newLine, True)

                    '' Save the new object to the database
                    acTrans.Commit()

                Else    ' only testing

                End If

                '' Dispose of the transaction
            End Using

        End Sub

        <CommandMethod("CCwindows")>
        Public Sub CCwindows() ' This method can have any name
            '' Get the editor, current document and database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                ' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                Dim blkRecId As ObjectId = ObjectId.Null

                If Not acBlkTbl.Has("Fortlaufnr") Then
                    ed.WriteMessage("Please use the casa clima template file!")

                    Exit Sub
                Else
                    blkRecId = acBlkTbl("Fortlaufnr")
                End If

                Dim pIntOpts As PromptIntegerOptions = New PromptIntegerOptions("")
                pIntOpts.Message = vbCrLf & "Start with index no. "
                '' Restrict input to positive and non-negative values
                pIntOpts.AllowZero = False
                pIntOpts.AllowNegative = False
                pIntOpts.AllowNone = False
                pIntOpts.UseDefaultValue = True
                pIntOpts.DefaultValue = 1
                Dim startNumber1 As PromptIntegerResult = acDoc.Editor.GetInteger(pIntOpts)
                Dim startNumber As Integer = startNumber1.Value

                ' Insert the block into the current space
                If blkRecId <> ObjectId.Null Then
                    Dim another As Boolean = True
                    While another = True
                        Dim pointOptions As New PromptPointOptions("Point for caption of new window: ")
                        pointOptions.AllowNone = True
                        Dim pointResult As PromptPointResult = ed.GetPoint(pointOptions)
                        If pointResult.Status = PromptStatus.Cancel Then
                            another = False
                        ElseIf pointResult.Status = PromptStatus.None Then
                            another = False
                        Else
                            Using acTrans1 As Transaction = acCurDb.TransactionManager.StartTransaction()
                                Dim selectedPoint As Point3d = pointResult.Value

                                Dim acBlkRef As New BlockReference(selectedPoint, blkRecId)
                                acBlkRef.ScaleFactors = New Scale3d(0.25)

                                Dim acCurSpaceBlkTblRec As BlockTableRecord

                                acCurSpaceBlkTblRec = acTrans1.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)

                                acCurSpaceBlkTblRec.AppendEntity(acBlkRef)
                                acTrans1.AddNewlyCreatedDBObject(acBlkRef, True)

                                ' add the attribute definitions.
                                Dim blkTblR As BlockTableRecord = blkRecId.GetObject(OpenMode.ForRead)
                                For Each objId As ObjectId In blkTblR
                                    Dim obj As DBObject = objId.GetObject(OpenMode.ForRead)
                                    If TypeOf obj Is AttributeDefinition Then
                                        Dim ad As AttributeDefinition = objId.GetObject(OpenMode.ForRead)
                                        Dim ar As AttributeReference = New AttributeReference()

                                        ar.SetAttributeFromBlock(ad, acBlkRef.BlockTransform)
                                        ar.Position = ad.Position.TransformBy(acBlkRef.BlockTransform)

                                        acBlkRef.AttributeCollection.AppendAttribute(ar)

                                        acTrans1.AddNewlyCreatedDBObject(ar, True)
                                    End If
                                Next

                                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection

                                For Each objId As ObjectId In attCol
                                    Dim att As AttributeReference = acTrans1.GetObject(objId, OpenMode.ForRead)
                                    If att.Tag = "P" Then
                                        att.UpgradeOpen()
                                        att.TextString = startNumber.ToString
                                    End If

                                Next

                                acTrans1.TransactionManager.QueueForGraphicsFlush()

                                acTrans1.Commit()
                            End Using

                            acTrans.TransactionManager.QueueForGraphicsFlush()

                            startNumber = startNumber + 1
                        End If

                    End While
                End If

                acTrans.Commit()
                acTrans.Dispose()
            End Using


        End Sub


        <CommandMethod("CCwalls", CommandFlags.Modal + CommandFlags.UsePickSet)>
        Public Sub CCwalls() ' This method can have any name
            '' Get the editor, current document and database
            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                '' Request for objects to be selected in the drawing area
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()

                '' If the prompt status is OK, objects were selected
                If acSSPrompt.Status = PromptStatus.OK Then
                    'open the current space (can be any BTR, usually model space)
                    Dim mSpace As BlockTableRecord = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)

                    Dim acSSet As SelectionSet = acSSPrompt.Value

                    Dim WallHeight As Double = 3.0
                    Dim WallHeight1 As Double = 4.0
                    Dim selectedPoint As Point3d
                    Dim countpoly As Integer = 0

                    Try
                        '' Step through the objects in the selection set
                        For Each acSSObj As SelectedObject In acSSet
                            '' Check to make sure a valid SelectedObject object was returned
                            If Not IsDBNull(acSSObj) Then
                                '' Open the selected object for write
                                Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)

                                If Not IsDBNull(acEnt) Then
                                    Dim aclayer As String = acEnt.Layer

                                    'ed.WriteMessage(acEnt.GetType.ToString())

                                    Dim polyline As Polyline = TryCast(acEnt, Polyline)
                                    If (polyline <> Nothing) Then
                                        If countpoly > 0 Then

                                        Else
                                            Dim pointOptions As New PromptPointOptions("Punkt für Abwicklung: ")
                                            Dim pointResult As PromptPointResult = ed.GetPoint(pointOptions)
                                            selectedPoint = pointResult.Value
                                        End If
                                        countpoly = countpoly + 1

                                        Dim maxSections As Integer
                                        If polyline.Closed Then
                                            maxSections = polyline.NumberOfVertices - 1
                                        Else
                                            maxSections = polyline.NumberOfVertices - 2
                                        End If

                                        For jj = 0 To maxSections

                                            Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                                            pKeyOpts.Message = vbLf & "Wall Type "
                                            pKeyOpts.Keywords.Add("01 Wand")
                                            pKeyOpts.Keywords.Add("02 Wand")
                                            pKeyOpts.Keywords.Add("03 Wand")
                                            pKeyOpts.Keywords.Add("04 Wand")
                                            pKeyOpts.Keywords.Add("05 Wand")
                                            pKeyOpts.Keywords.Add("06 Wand")
                                            pKeyOpts.Keywords.Add("07 Wand")
                                            pKeyOpts.Keywords.Add("08 Wand")
                                            pKeyOpts.Keywords.Add("09 Wand")
                                            pKeyOpts.Keywords.Add("10 Wand")
                                            pKeyOpts.Keywords.Add("11 Wand")
                                            pKeyOpts.Keywords.Add("12 Wand")
                                            pKeyOpts.Keywords.Add("13 Wand")
                                            pKeyOpts.Keywords.Add("14 Wand")
                                            pKeyOpts.Keywords.Add("15 Wand")
                                            Try
                                                pKeyOpts.Keywords.Default = aclayer
                                            Catch ex As Exception
                                                pKeyOpts.Keywords.Default = "01 Wand"
                                            End Try
                                            pKeyOpts.AllowNone = False

                                            Dim pKeyRes As PromptResult = acDoc.Editor.GetKeywords(pKeyOpts)

                                            Dim WallType As String = "Wall"

                                            Dim pDblOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf &
                                                                    "Give (left) wall height [m]: ")
                                            pDblOpts.AllowZero = True
                                            pDblOpts.AllowNegative = False
                                            pDblOpts.DefaultValue = WallHeight
                                            pDblOpts.Keywords.Add("Wall")
                                            pDblOpts.Keywords.Add("Trapez")
                                            pDblOpts.Keywords.Add("DoubleTrapez")
                                            pDblOpts.Keywords.Default = "Wall"

                                            Dim readinput As Boolean = True
                                            Do While readinput
                                                Dim pDblRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDblOpts)
                                                If pDblRes.Status = PromptStatus.Keyword Then
                                                    WallType = pDblRes.StringResult
                                                Else
                                                    WallHeight = pDblRes.Value
                                                    readinput = False
                                                End If
                                            Loop

                                            Dim newSegment As LineSegment3d = polyline.GetLineSegmentAt(jj)

                                            Dim newlayer As String = aclayer
                                            If pKeyRes.StringResult.Length <= 3 Then
                                                newlayer = pKeyRes.StringResult & " Wand"
                                            Else
                                                newlayer = pKeyRes.StringResult
                                            End If

                                            Using acLine As Line = New Line()
                                                acLine.StartPoint = newSegment.StartPoint
                                                acLine.EndPoint = newSegment.EndPoint

                                                acLine.Layer = newlayer

                                                Dim acLineAngle As Double = acLine.Angle / (2 * Math.PI) * 360
                                                acLineAngle = acLineAngle + 180
                                                If acLineAngle > 360 Then
                                                    acLineAngle = acLineAngle - 360
                                                End If

                                                mSpace.AppendEntity(acLine)
                                                acTrans.AddNewlyCreatedDBObject(acLine, True)
                                                acTrans.TransactionManager.QueueForGraphicsFlush()

                                                Dim CaptionOrientation As String = ""
                                                Select Case acLineAngle
                                                    Case 0 To (45 - 45 / 2)
                                                        CaptionOrientation = "S"
                                                    Case (45 - 45 / 2) To (90 - 45 / 2)
                                                        CaptionOrientation = "SO"
                                                    Case (90 - 45 / 2) To (135 - 45 / 2)
                                                        CaptionOrientation = "O"
                                                    Case (135 - 45 / 2) To (180 - 45 / 2)
                                                        CaptionOrientation = "NO"
                                                    Case (180 - 45 / 2) To (225 - 45 / 2)
                                                        CaptionOrientation = "N"
                                                    Case (225 - 45 / 2) To (270 - 45 / 2)
                                                        CaptionOrientation = "NW"
                                                    Case (270 - 45 / 2) To (315 - 45 / 2)
                                                        CaptionOrientation = "W"
                                                    Case (315 - 45 / 2) To (360 - 45 / 2)
                                                        CaptionOrientation = "SW"
                                                    Case (360 - 45 / 2) To 360
                                                        CaptionOrientation = "S"
                                                End Select

                                                'Using objText As MText = New MText
                                                '    '' Specify the insertion point of the MText object
                                                '    Dim extents As Extents3d = acLine.GeometricExtents
                                                '    Dim acLineMidpoint As Point3d = extents.MinPoint + (extents.MaxPoint - extents.MinPoint) / 2.0

                                                '    objText.Location = acLineMidpoint

                                                '    '' Set the text string for the MText object
                                                '    objText.Contents = CaptionOrientation

                                                '    '' Set the text style for the MText object
                                                '    objText.TextStyleId = acCurDb.Textstyle
                                                '    objText.TextHeight = 0.25
                                                '    objText.Rotation = acLineAngle * (2 * Math.PI) / 360
                                                '    objText.Height = 0.5
                                                '    objText.Layer = newlayer
                                                '    Try
                                                '        objText.Attachment = AttachmentPoint.BottomCenter
                                                '    Catch ex As Exception

                                                '    End Try

                                                '    '' Appends the new MText object to model space
                                                '    mSpace.AppendEntity(objText)

                                                '    '' Appends to new MText object to the active transaction
                                                '    acTrans.AddNewlyCreatedDBObject(objText, True)
                                                'End Using

                                                Using acPoly As Polyline = New Polyline()
                                                    Using objText As MText = New MText
                                                        '' Specify the insertion point of the MText object
                                                        Dim extents1 As Extents3d = acLine.GeometricExtents
                                                        Dim acLineMidpoint As Point3d

                                                        acLineMidpoint = New Point3d(selectedPoint.Coordinate(0) + newSegment.Length / 2.0, selectedPoint.Coordinate(1) - 0.5, 0)

                                                        Select Case WallType
                                                            Case "Wall"
                                                                acPoly.AddVertexAt(0, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1)), 0, 0, 0)
                                                                acPoly.AddVertexAt(1, New Point2d(selectedPoint.Coordinate(0) + newSegment.Length, selectedPoint.Coordinate(1)), 0, 0, 0)
                                                                acPoly.AddVertexAt(2, New Point2d(selectedPoint.Coordinate(0) + newSegment.Length, selectedPoint.Coordinate(1) + WallHeight), 0, 0, 0)
                                                                acPoly.AddVertexAt(3, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1) + WallHeight), 0, 0, 0)
                                                                acPoly.AddVertexAt(4, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1)), 0, 0, 0)
                                                            Case "Trapez"
                                                                Dim pDblOpts1 As PromptDoubleOptions = New PromptDoubleOptions(vbLf &
                                                                        "Give second (right) wall height [m]: ")
                                                                pDblOpts1.AllowZero = True
                                                                pDblOpts1.AllowNegative = True
                                                                pDblOpts1.DefaultValue = WallHeight1
                                                                Dim pDblRes1 As PromptDoubleResult = acDoc.Editor.GetDouble(pDblOpts1)
                                                                WallHeight1 = pDblRes1.Value

                                                                acPoly.AddVertexAt(0, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1)), 0, 0, 0)
                                                                acPoly.AddVertexAt(1, New Point2d(selectedPoint.Coordinate(0) + newSegment.Length, selectedPoint.Coordinate(1)), 0, 0, 0)
                                                                acPoly.AddVertexAt(2, New Point2d(selectedPoint.Coordinate(0) + newSegment.Length, selectedPoint.Coordinate(1) + WallHeight1), 0, 0, 0)
                                                                acPoly.AddVertexAt(3, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1) + WallHeight), 0, 0, 0)
                                                                acPoly.AddVertexAt(4, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1)), 0, 0, 0)

                                                                WallHeight = WallHeight1
                                                            Case "DoubleTrapez"
                                                                Dim pDblOpts1 As PromptDoubleOptions = New PromptDoubleOptions(vbLf &
                                                                        "Give second (middle) wall height [m]: ")
                                                                pDblOpts1.AllowZero = True
                                                                pDblOpts1.AllowNegative = True
                                                                pDblOpts1.DefaultValue = WallHeight1
                                                                Dim pDblRes1 As PromptDoubleResult = acDoc.Editor.GetDouble(pDblOpts1)
                                                                WallHeight1 = pDblRes1.Value

                                                                Dim pDblOpts3 As PromptDoubleOptions = New PromptDoubleOptions(vbLf &
                                                                        "Give distance to left wall point [m]: ")
                                                                pDblOpts3.AllowZero = True
                                                                pDblOpts3.AllowNegative = True
                                                                pDblOpts3.DefaultValue = WallHeight
                                                                Dim pDblRes3 As PromptDoubleResult = acDoc.Editor.GetDouble(pDblOpts3)
                                                                Dim WallWidth As Double = pDblRes3.Value

                                                                Dim pDblOpts2 As PromptDoubleOptions = New PromptDoubleOptions(vbLf &
                                                                        "Give third (right) wall height [m]: ")
                                                                pDblOpts2.AllowZero = True
                                                                pDblOpts2.AllowNegative = True
                                                                pDblOpts2.DefaultValue = WallHeight
                                                                Dim pDblRes2 As PromptDoubleResult = acDoc.Editor.GetDouble(pDblOpts2)
                                                                Dim WallHeight2 As Double = pDblRes2.Value

                                                                acPoly.AddVertexAt(0, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1)), 0, 0, 0)
                                                                acPoly.AddVertexAt(1, New Point2d(selectedPoint.Coordinate(0) + newSegment.Length, selectedPoint.Coordinate(1)), 0, 0, 0)
                                                                acPoly.AddVertexAt(2, New Point2d(selectedPoint.Coordinate(0) + newSegment.Length, selectedPoint.Coordinate(1) + WallHeight2), 0, 0, 0)
                                                                acPoly.AddVertexAt(3, New Point2d(selectedPoint.Coordinate(0) + WallWidth, selectedPoint.Coordinate(1) + WallHeight1), 0, 0, 0)
                                                                acPoly.AddVertexAt(4, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1) + WallHeight), 0, 0, 0)
                                                                acPoly.AddVertexAt(5, New Point2d(selectedPoint.Coordinate(0), selectedPoint.Coordinate(1)), 0, 0, 0)

                                                                WallHeight = WallHeight2
                                                        End Select

                                                        objText.Location = acLineMidpoint
                                                        objText.Contents = CaptionOrientation
                                                        objText.TextStyleId = acCurDb.Textstyle
                                                        objText.TextHeight = 0.25
                                                        objText.Height = 0.5
                                                        objText.Layer = newlayer
                                                        objText.Attachment = AttachmentPoint.BottomCenter
                                                        mSpace.AppendEntity(objText)
                                                        acTrans.AddNewlyCreatedDBObject(objText, True)
                                                    End Using

                                                    selectedPoint = New Point3d(selectedPoint.Coordinate(0) + newSegment.Length, selectedPoint.Coordinate(1), selectedPoint.Coordinate(2))

                                                    acPoly.Layer = newlayer

                                                    '' Add the new object to the block table record and the transaction
                                                    mSpace.AppendEntity(acPoly)
                                                    acTrans.AddNewlyCreatedDBObject(acPoly, True)
                                                    acTrans.TransactionManager.QueueForGraphicsFlush()

                                                    Dim extents As Extents3d = acPoly.GeometricExtents
                                                    Dim selectedPoint1 As Point3d = extents.MinPoint + (extents.MaxPoint - extents.MinPoint) / 2.0

                                                    Dim area As Double = acPoly.Area

                                                    '' Adds all to an object id array
                                                    Dim acObjIdColl As ObjectIdCollection = New ObjectIdCollection()
                                                    acObjIdColl.Add(acPoly.ObjectId)

                                                    Using objText As MText = New MText
                                                        '' Specify the insertion point of the MText object
                                                        objText.Location = selectedPoint1

                                                        '' Set the text string for the MText object
                                                        objText.Contents = Math.Round(area, 2).ToString & " qm"

                                                        '' Set the text style for the MText object
                                                        objText.TextStyleId = acCurDb.Textstyle
                                                        objText.TextHeight = 0.25
                                                        objText.Layer = acPoly.Layer
                                                        Try
                                                            objText.Attachment = AttachmentPoint.MiddleCenter
                                                        Catch ex As Exception

                                                        End Try

                                                        '' Appends the new MText object to model space
                                                        mSpace.AppendEntity(objText)

                                                        '' Appends to new MText object to the active transaction
                                                        acTrans.AddNewlyCreatedDBObject(objText, True)

                                                        Dim acObjIdCol2 As ObjectIdCollection = New ObjectIdCollection()
                                                        acObjIdCol2.Add(objText.ObjectId)

                                                        Using objHatch As Hatch = New Hatch
                                                            mSpace.AppendEntity(objHatch)
                                                            objHatch.PatternScale = 0.1
                                                            objHatch.Layer = acPoly.Layer
                                                            acTrans.AddNewlyCreatedDBObject(objHatch, True)

                                                            '' Set the properties of the hatch object
                                                            '' Associative must be set after the hatch object is appended to the 
                                                            '' block table record and before AppendLoop
                                                            objHatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31")
                                                            objHatch.Associative = True
                                                            objHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl)
                                                            objHatch.AppendLoop(HatchLoopTypes.TextIsland, acObjIdCol2)
                                                            objHatch.EvaluateHatch(True)
                                                        End Using
                                                    End Using

                                                    acTrans.TransactionManager.QueueForGraphicsFlush()
                                                End Using

                                            End Using

                                            acTrans.TransactionManager.QueueForGraphicsFlush()
                                        Next

                                        polyline.Erase()
                                        acTrans.TransactionManager.QueueForGraphicsFlush()

                                    Else
                                        ed.WriteMessage("Selected object not a polyline!")
                                    End If

                                End If
                            End If
                        Next

                        'Save the new object to the database
                        acTrans.Commit()
                    Catch ex As Exception

                    End Try
                Else    ' only testing

                End If

                '' Dispose of the transaction
            End Using

        End Sub

        ' declare a paletteset object, this will only be created once
        Public myPaletteSet As Autodesk.AutoCAD.Windows.PaletteSet
        ' we need a palette which will be housed by the paletteSet
        Public myPalette As UserControl1

        ' palette command
        <CommandMethod("palette")>
        Public Sub palette()
            ' check to see if it is valid
            If (myPaletteSet = Nothing) Then
                ' create a new palette set, with a unique guid
                myPaletteSet = New Autodesk.AutoCAD.Windows.PaletteSet("My Palette", New Guid("D61D0875-A507-4b73-8B5F-9266BEACD596"))
                ' now create a palette inside, this has our tree control
                myPalette = New UserControl1
                ' now add the palette to the paletteset
                myPaletteSet.Add("Palette1", myPalette)

            End If

            ' now display the paletteset
            myPaletteSet.Visible = True


        End Sub

        <CommandMethod("addDBEvents")>
        Public Sub addDBEvents()
            ' the palette needs to be created for this
            If myPalette Is Nothing Then
                ' get the editor object
                Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                ' now write to the command line
                ed.WriteMessage(vbCr + "Please call the 'palette' command first")
                Exit Sub
            End If

            ' get the current working database
            Dim curDwg As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
            ' add a handlers for what we need
            AddHandler curDwg.ObjectAppended, New ObjectEventHandler(AddressOf callback_ObjectAppended)
            AddHandler curDwg.ObjectErased, New ObjectErasedEventHandler(AddressOf callback_ObjectErased)
            AddHandler curDwg.ObjectReappended, New ObjectEventHandler(AddressOf callback_ObjectReappended)
            AddHandler curDwg.ObjectUnappended, New ObjectEventHandler(AddressOf callback_ObjectUnappended)

        End Sub

        Private Sub callback_ObjectAppended(ByVal sender As Object, ByVal e As ObjectEventArgs)

            ' add the class name of the object to the tree view
            Dim newNode As System.Windows.Forms.TreeNode = myPalette.TreeView1.Nodes.Add(e.DBObject.GetType().ToString())
            ' we need to record its id for recognition later
            newNode.Tag = e.DBObject.ObjectId.ToString()

        End Sub

        Private Sub callback_ObjectErased(ByVal sender As Object, ByVal e As ObjectErasedEventArgs)

            ' if the object was erased
            If e.Erased Then
                ' find the object in the treeview control so we can remove it
                For Each node As Forms.TreeNode In myPalette.TreeView1.Nodes
                    ' is this the one we want
                    If (node.Tag = e.DBObject.ObjectId.ToString) Then
                        node.Remove()
                        Exit For
                    End If
                Next
                ' if the object was unerased
            Else
                ' add the class name of the object to the tree view
                Dim newNode As System.Windows.Forms.TreeNode = myPalette.TreeView1.Nodes.Add(e.DBObject.GetType().ToString())
                ' we need to record its id for recognition later
                newNode.Tag = e.DBObject.ObjectId.ToString()
            End If
        End Sub

        Private Sub callback_ObjectReappended(ByVal sender As Object, ByVal e As ObjectEventArgs)

            ' add the class name of the object to the tree view
            Dim newNode As Forms.TreeNode = myPalette.TreeView1.Nodes.Add(e.DBObject.GetType().ToString())
            ' we need to record its id for recognition later
            newNode.Tag = e.DBObject.ObjectId.ToString()

        End Sub

        Private Sub callback_ObjectUnappended(ByVal sender As Object, ByVal e As ObjectEventArgs)

            ' find the object in the treeview control so we can remove it
            For Each node As Forms.TreeNode In myPalette.TreeView1.Nodes
                ' is this the one we want
                If (node.Tag = e.DBObject.ObjectId.ToString) Then
                    node.Remove()
                    Exit For
                End If
            Next

        End Sub

    End Class

End Namespace