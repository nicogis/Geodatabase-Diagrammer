
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
' Imports Microsoft.Office.Interop


Module VisioDiagram

#Region "Member Variables"

    Const visSelect% = 2

    ' 2.0 (use early binding)   (do not dim as Object)
    ' http://support.microsoft.com/kb/309603
    Dim appVisio As Microsoft.Office.Interop.Visio.Application ' Instance of Visio.
    Dim docsObj As Microsoft.Office.Interop.Visio.Documents 'Documents collection of instance
    Dim docObj As Microsoft.Office.Interop.Visio.Document ' Document to work in.
    Dim stnObj As Microsoft.Office.Interop.Visio.Document    'Stencil that contains master
    Dim mastObj As Microsoft.Office.Interop.Visio.Master    'Master to drop
    Dim pagsObj As Microsoft.Office.Interop.Visio.Pages    'Pages collection of document
    Dim pagObj As Microsoft.Office.Interop.Visio.Page    'Page to work in
    Dim shpObj As Microsoft.Office.Interop.Visio.Shape      'Instance of master on page

    Dim selObj As Object     'Selection object
    Dim numPages As Integer  'Number of pages in document

    Dim curX As Double       'Used for placing objects on Visio page
    Dim curY As Double
    Dim curHeight As Double
    Dim locX As Double, locY As Double

    Dim fieldWidth_FieldName As Double
    Dim fieldWidth_DataType As Double
    Dim fieldWidth_AllowNulls As Double
    Dim fieldWidth_DefaultValue As Double
    Dim fieldWidth_Domain As Double
    Dim fieldWidth_Precision As Double
    Dim fieldWidth_Scale As Double
    Dim fieldWidth_Length As Double

    Dim boxHeight As Double

    Dim curClassType As String

    Dim cVisioFileName As String

    Dim cNames As String
    Dim lFirstSubtype As Boolean

    ' public variables that are used to capture form settings.
    Public Structure DiagrammerSettings

        Dim Postcript As Boolean
        Dim TrueType As Boolean

        Dim Summary As Boolean
        Dim UseAbstract As Boolean
        Dim FieldMetadata As Boolean
        Dim FieldAlias As Boolean
        Dim OmitAnno As Boolean

    End Structure



#End Region

#Region "Public Functions"

    Public Function StartDiagram(ByVal cPathName As String, ByVal strVisioOutputFile As String)

        Try
            '  Dim lenPath As Long
            '  lenPath = Len(cPathName)
            '  cVisioFileName = Left(cPathName, lenPath - 3) & "vsd"

            cVisioFileName = strVisioOutputFile

            OpenVisioDrawing(cVisioFileName, cPathName)

        Catch ex As Exception
            ExHandle(ex)
        End Try
#Disable Warning BC42105 ' Function 'StartDiagram' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'StartDiagram' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Function OpenVisioDrawing(ByVal fileName As String, ByVal pathName As String) As Boolean
        'We open the visio drawing by starting Visio, opening the document, and setting
        'the first page of the document.
        Dim result As Boolean = False
        Try


            appVisio = CreateObject("visio.InvisibleApp") '("visio.application")
            appVisio.AlertResponse = 1 'This answers "Ok" to model dialogs
            docsObj = appVisio.Documents



            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings
            If dSets.Postcript Then
                docObj = docsObj.Add("GeodatabaseDiagrammerPS.vst")
            End If
            If dSets.TrueType Then
                docObj = docsObj.Add("GeodatabaseDiagrammerTT.vst")
            End If
            'If pFormGDBDiagrammer.optPS.Value = True Then docObj = docsObj.Add("GeodatabaseDiagrammerPS.vst")
            'If FormGDBDiagrammer.optTT.Value = True Then docObj = docsObj.Add("GeodatabaseDiagrammerTT.vst")

            pagsObj = docObj.Pages
            pagObj = pagsObj.Item(1)
            numPages = 1

            'Set stencil
            If dSets.Postcript Then
                stnObj = docsObj.Open("GeodatabaseDiagrammerPS.vss")
            End If
            If dSets.TrueType Then
                stnObj = docsObj.Open("GeodatabaseDiagrammerTT.vss")
            End If
            'If FormGDBDiagrammer.optPS.Value = True Then stnObj = docsObj.Open("GeodatabaseDiagrammerPS.vss")
            'If FormGDBDiagrammer.optTT.Value = True Then stnObj = docsObj.Open("GeodatabaseDiagrammerTT.vss")

            curX = 0
            curY = 0

            fieldWidth_FieldName = 1.125
            'fieldWidth_DataType = 0.5
            fieldWidth_DataType = 0.5625
            fieldWidth_AllowNulls = 0.28125
            'fieldWidth_DefaultValue = 0.28125
            fieldWidth_DefaultValue = 1
            'fieldWidth_Domain = 0.96875
            fieldWidth_Domain = 1
            fieldWidth_Precision = 0.28125
            fieldWidth_Scale = 0.28125
            fieldWidth_Length = 0.28125

            mastObj = stnObj.Masters("GDB Name")
            shpObj = pagObj.Drop(mastObj, 2, 33.5)
            shpObj.Text = pathName
            shpObj = pagObj.Drop(mastObj, 2, 33.25)

            shpObj.Text = FormatDateTime(Now, DateFormat.LongDate)
            'shpObj.Text = Format(Date, "Long Date")


        Catch ex As Exception
            ExHandle(ex)
        End Try
        Return result
    End Function

    Public Function SectionHeader()
        Try

            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            If dSets.Summary Then
                mastObj = stnObj.Masters("Summary Title")
            Else
                mastObj = stnObj.Masters("Detail Title")
            End If

            Call positionOnPage()
            curY = curY - 0.5
            shpObj = pagObj.Drop(mastObj, curX, curY)

            curY = curY - 1.0#

        Catch ex As Exception
            ExHandle(ex)
        End Try

#Disable Warning BC42105 ' Function 'SectionHeader' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'SectionHeader' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Function StartFeatureClass(ByVal FeatureClassName As String, ByVal FeatureType As Long, ByVal ShapeType As Long, ByVal numFields As Long, ByVal sHasMs As String, ByVal sHasZs As String, ByVal sAbstract As String)

        Dim shapeDescription As String

        Try

            Call positionOnPage()

            locX = curX
            locY = curY
            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            If dSets.Summary Then

                'First, drop little version of feature class diagram
                Select Case FeatureType

                    Case esriFeatureType.esriFTSimple

                        Select Case ShapeType

                            Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint
                                mastObj = stnObj.Masters("SmallPoint")

                            Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryMultipoint
                                mastObj = stnObj.Masters("SmallPoint")

                            Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline
                                mastObj = stnObj.Masters("SmallLine")

                            Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon
                                mastObj = stnObj.Masters("SmallPolygon")

                            Case Else
                                MsgBox("In modVisioDiagram.StartFeatureClass, unknown shape type")

                        End Select

                    Case esriFeatureType.esriFTComplexEdge
                        mastObj = stnObj.Masters("SmallComplexEdge")

                    Case esriFeatureType.esriFTComplexJunction
                        mastObj = stnObj.Masters("SmallComplexJunction")

                    Case esriFeatureType.esriFTSimpleJunction
                        mastObj = stnObj.Masters("SmallSimpleJunction")

                    Case esriFeatureType.esriFTSimpleEdge
                        mastObj = stnObj.Masters("SmallSimpleEdge")

                    Case esriFeatureType.esriFTAnnotation
                        mastObj = stnObj.Masters("SmallAnnotation")

                    Case esriFeatureType.esriFTDimension
                        mastObj = stnObj.Masters("SmallDimension")

                End Select

                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Ungroup()
                mastObj = stnObj.Masters("TableClassName")
                shpObj = pagObj.Drop(mastObj, locX + 0.0313, locY - 0.1563)
                shpObj.Text = FeatureClassName

                locY = locY - 0.5
                curY = curY - 0.5

            Else

                curClassType = FeatureType

                'Calculate and set height of box
                boxHeight = 0.5625 + (numFields * 0.125)

                'Drop background
                mastObj = stnObj.Masters("Class background")
                shpObj = pagObj.Drop(mastObj, locX + 0.0313, locY - 0.0313)

                shpObj.Cells("Height").ResultIU = boxHeight
                shpObj.Cells("Width").ResultIU = 4.8125


                'Drop feature class box
                mastObj = stnObj.Masters("Feature class box")
                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Cells("Height").ResultIU = boxHeight


                Select Case FeatureType

                    Case esriFeatureType.esriFTSimple
                        shpObj.Text = "Simple feature class"

                    Case esriFeatureType.esriFTComplexEdge
                        shpObj.Text = "Complex edge feature class"

                    Case esriFeatureType.esriFTComplexJunction
                        shpObj.Text = "Complex junction feature class"

                    Case esriFeatureType.esriFTSimpleJunction
                        shpObj.Text = "Simple junction feature class"

                    Case esriFeatureType.esriFTSimpleEdge
                        shpObj.Text = "Simple edge feature class"

                    Case esriFeatureType.esriFTAnnotation
                        shpObj.Text = "Annotation feature class"

                    Case esriFeatureType.esriFTDimension
                        shpObj.Text = "Dimension feature class"

                End Select

                'Drop feature class type icon
                If FeatureType = esriFeatureType.esriFTAnnotation Then
                    mastObj = stnObj.Masters("Annotation icon")

                ElseIf FeatureType = esriFeatureType.esriFTDimension Then
                    mastObj = stnObj.Masters("Dimension icon")

                Else

                    Select Case ShapeType

                        Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint
                            mastObj = stnObj.Masters("Point icon")
                            shapeDescription = "Point"

                        Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryMultipoint
                            mastObj = stnObj.Masters("Point icon")
                            shapeDescription = "Multipoint"

                        Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline
                            mastObj = stnObj.Masters("Polyline icon")
                            shapeDescription = "Polyline"

                        Case ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon
                            mastObj = stnObj.Masters("Polygon icon")
                            shapeDescription = "Polygon"

                        Case Else
                            MsgBox("In modVisioDiagram.StartFeatureClass, unknown shape type")

                    End Select

                End If

                shpObj = pagObj.Drop(mastObj, locX + 0.06, locY - 0.06)

                'Drop class name
                mastObj = stnObj.Masters("Class name")
                shpObj = pagObj.Drop(mastObj, locX + 0.2813, locY - 0.1563)
                shpObj.Text = FeatureClassName

                'Drop contains statements
                mastObj = stnObj.Masters("Contains")
                shpObj = pagObj.Drop(mastObj, locX + 3.5938, locY - 0.0156)

                mastObj = stnObj.Masters("ContainsNo")
                shpObj = pagObj.Drop(mastObj, locX + 4.3438, locY - 0.0156)
#Disable Warning BC42104 ' Variable 'shapeDescription' is used before it has been assigned a value. A null reference exception could result at runtime.
                shpObj.Text = shapeDescription
#Enable Warning BC42104 ' Variable 'shapeDescription' is used before it has been assigned a value. A null reference exception could result at runtime.

                mastObj = stnObj.Masters("ContainsNo")
                shpObj = pagObj.Drop(mastObj, locX + 4.3438, locY - 0.1094)
                shpObj.Text = sHasMs

                mastObj = stnObj.Masters("ContainsNo")
                shpObj = pagObj.Drop(mastObj, locX + 4.3438, locY - 0.2031)
                shpObj.Text = sHasZs

                'Drop field header
                mastObj = stnObj.Masters("Class header")
                shpObj = pagObj.Drop(mastObj, locX, locY - 0.3125)
                shpObj.Ungroup()


                'Drop table description placeholder text
                mastObj = stnObj.Masters("Table description")
                shpObj = pagObj.Drop(mastObj, locX + 4.9063, locY)

                ' Change text to metadata abstract
                If dSets.UseAbstract Then
                    shpObj.Text = sAbstract
                End If

            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try

#Disable Warning BC42105 ' Function 'StartFeatureClass' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'StartFeatureClass' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Sub startTable(ByVal tableName As String, ByVal numFields As Long, ByVal sAbstract As String)
        ' note: was a function
        Try
            curClassType = "Table"

            Call positionOnPage()

            locX = curX
            locY = curY

            'If summary checked, write out summary table shape...

            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            If dSets.Summary Then
                mastObj = stnObj.Masters("SmallTable")
                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Ungroup()
                mastObj = stnObj.Masters("TableClassName")
                shpObj = pagObj.Drop(mastObj, locX + 0.0313, locY - 0.1563)
                shpObj.Text = tableName

                locY = locY - 0.5
                curY = curY - 0.5

                GoTo endOfFunction
            End If

            'Calculate and set height of box
            boxHeight = 0.5625 + (numFields * 0.125)

            'Drop background
            mastObj = stnObj.Masters("Class background")
            shpObj = pagObj.Drop(mastObj, locX + 0.0313, locY - 0.0313)
            shpObj.Cells("Height").ResultIU = boxHeight
            shpObj.Cells("Width").ResultIU = 4.8125

            'Drop feature class box
            mastObj = stnObj.Masters("Feature class box")
            shpObj = pagObj.Drop(mastObj, locX, locY)
            shpObj.Cells("Height").ResultIU = boxHeight

            shpObj.Text = "Table"

            mastObj = stnObj.Masters("Table icon")
            shpObj = pagObj.Drop(mastObj, locX + 0.06, locY - 0.06)

            'Drop class name
            mastObj = stnObj.Masters("Class name")
            shpObj = pagObj.Drop(mastObj, locX + 0.2813, locY - 0.1563)
            shpObj.Text = tableName

            'Drop field header
            mastObj = stnObj.Masters("Class header")
            shpObj = pagObj.Drop(mastObj, locX, locY - 0.3125)
            shpObj.Ungroup()

            'Drop table description placeholder text
            mastObj = stnObj.Masters("Table description")
            shpObj = pagObj.Drop(mastObj, locX + 4.9063, locY)

            ' Use abstract text instead
            If dSets.UseAbstract Then
                shpObj.Text = sAbstract
            End If

endOfFunction:

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Public Sub DiagramFeatureClass(ByRef pFields As IFields, ByRef pDatasetName As String, Optional ByVal metadata As IPropertySet = Nothing)
        ' note: was a function
        Try

            Dim curDomain As IDomain

            Dim fieldX As Double
            Dim lName As Boolean, lType As Boolean, lAllowNulls As Boolean, lDefaultValue As Boolean, lDomain As Boolean, lPrecision As Boolean, lScale As Boolean, lLength As Boolean
            Dim cFieldType As String
            Dim pFieldCounter As Long
            Dim cNameFieldType As String

            'Dim cType As String
            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            If dSets.Summary Then GoTo endOfFunction

            locY = locY - 0.4375

            For pFieldCounter = 0 To pFields.FieldCount - 1 Step 1

                'This is the X value for the current field
                fieldX = locX
                locY = locY - 0.125

                'Initialize booleans that indicate whether to drop a value or null box for the current field property
                lName = True
                lType = True
                lAllowNulls = True
                lDefaultValue = True
                lDomain = True
                lPrecision = True
                lScale = True
                lLength = True

                Select Case pFields.Field(pFieldCounter).Type

                    Case esriFieldType.esriFieldTypeOID
                        lAllowNulls = False
                        lDefaultValue = False
                        lDomain = False
                        lPrecision = False
                        lScale = False
                        lLength = False

                    Case esriFieldType.esriFieldTypeGeometry
                        lDefaultValue = False
                        lDomain = False
                        lPrecision = False
                        lScale = False
                        lLength = False

                    Case esriFieldType.esriFieldTypeSmallInteger
                        lScale = False
                        lLength = False

                    Case esriFieldType.esriFieldTypeInteger
                        lScale = False
                        lLength = False

                    Case esriFieldType.esriFieldTypeSingle
                        lLength = False

                    Case esriFieldType.esriFieldTypeDouble
                        lLength = False

                    Case esriFieldType.esriFieldTypeString
                        lPrecision = False
                        lScale = False

                End Select

                'Now start adding boxes for each field property

                'Add field name box
                cFieldType = setFieldType("Name", pFields.Field(pFieldCounter).Name, lName)
                cNameFieldType = cFieldType
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_FieldName
                shpObj.Text = pFields.Field(pFieldCounter).Name
                If lName = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_FieldName

                'Add field type box
                cFieldType = setFieldType("Type", pFields.Field(pFieldCounter).Name, lType)
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, fieldX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_DataType
                shpObj.Text = findFieldType(pFields.Field(pFieldCounter).Type)
                If lType = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_DataType

                'Add allow nulls box
                cFieldType = setFieldType("AllowNulls", pFields.Field(pFieldCounter).Name, lAllowNulls)
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, fieldX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_AllowNulls
                If lAllowNulls Then
                    If pFields.Field(pFieldCounter).IsNullable Then
                        shpObj.Text = "Yes"
                    Else
                        shpObj.Text = "No"
                    End If
                End If
                If lAllowNulls = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_AllowNulls

                'Add default value box
                cFieldType = setFieldType("DefaultValue", pFields.Field(pFieldCounter).Name, lDefaultValue)
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, fieldX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_DefaultValue

                If Not TypeOf pFields.Field(pFieldCounter).DefaultValue Is DBNull Then
                    shpObj.Text = pFields.Field(pFieldCounter).DefaultValue
                End If

                'shpObj.Text = CStr(pFields.Field(pFieldCounter).DefaultValue)
                If lDefaultValue = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_DefaultValue

                'Add domain box
                cFieldType = setFieldType("Domain", pFields.Field(pFieldCounter).Name, lDomain)
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, fieldX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_Domain

                curDomain = pFields.Field(pFieldCounter).Domain
                If Not curDomain Is Nothing Then
                    shpObj.Text = pFields.Field(pFieldCounter).Domain.Name
                End If
                If lDomain = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_Domain

                'Add precision box
                cFieldType = setFieldType("Precision", pFields.Field(pFieldCounter).Name, lPrecision)
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, fieldX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_Precision

                shpObj.Text = CStr(pFields.Field(pFieldCounter).Precision)
                If lPrecision = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_Precision

                'Add scale box
                cFieldType = setFieldType("Scale", pFields.Field(pFieldCounter).Name, lScale)
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, fieldX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_Scale

                shpObj.Text = CStr(pFields.Field(pFieldCounter).Scale)
                If lScale = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_Scale

                'Add length box
                cFieldType = setFieldType("Length", pFields.Field(pFieldCounter).Name, lLength)
                mastObj = stnObj.Masters(cFieldType)
                shpObj = pagObj.Drop(mastObj, fieldX, locY)
                shpObj.Cells("Width").ResultIU = fieldWidth_Length

                shpObj.Text = CStr(pFields.Field(pFieldCounter).Length)
                If lLength = False Then shpObj.Text = " "
                fieldX = fieldX + fieldWidth_Scale

                'Add field description text (only if the current field is not a predefined (reserved) field)
                If cNameFieldType = "Field box" Then

                    mastObj = stnObj.Masters("Field description")
                    shpObj = pagObj.Drop(mastObj, fieldX + 0.0938, locY)


                    ' Option to use the field alias instead of the generic "Place a succinct ...." text.
                    If dSets.FieldAlias Then
                        shpObj.Text = CStr(pFields.Field(pFieldCounter).AliasName)
                    End If

                    ' Option to use the field metadata 'definition' instead of the generic "Place a succinct ...." text.  MC 12/5/07
                    If dSets.FieldMetadata Then
                        If Not metadata Is Nothing Then
                            Dim vDef As Object
                            'we need the attr node that has the attrlabl subnode that matches the columnname
                            vDef = metadata.GetProperty("eainfo/detailed/attr[attrlabl = '" & pFields.Field(pFieldCounter).Name & "']/attrdef")
                            If (Not vDef Is Nothing) AndAlso (Not IsDBNull(vDef)) Then
                                shpObj.Text = CStr(vDef(0))
                            End If
                            ' OLD:
                            'If ((Not IsNull(vDef)) And (Not IsEmpty(vDef))) Then
                            '    shpObj.Text = CStr(vDef(0))
                            'End If
                        End If
                    End If

                End If

            Next

            'drop outline box
            mastObj = stnObj.Masters("Box frame")
            shpObj = pagObj.Drop(mastObj, curX, curY)
            shpObj.Cells("Height").ResultIU = boxHeight
            shpObj.Cells("Width").ResultIU = 4.8125

            curY = locY - 0.5

endOfFunction:

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Public Sub startSubtype(ByVal sClassName As String, ByVal sSubtypeField As String, ByVal lSubtypeFromBottom As Long, ByVal sDefaultSubtype As String, ByVal numSubtypes As Long, ByVal numUniqueValues As Long)
        ' note: was a function
        Try
            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            If dSets.Summary Then 'FormGDBDiagrammer.chkSummary.Value = 1 Then
                'Drop subtypes background
                mastObj = stnObj.Masters("SmallSubtype")
                shpObj = pagObj.Drop(mastObj, curX, curY + 0.1875)
                cNames = ""
                lFirstSubtype = True
                GoTo endOfFunction
            End If

            locX = curX
            locY = curY

            'Drop subtype header
            mastObj = stnObj.Masters("Subtype header")
            shpObj = pagObj.Drop(mastObj, locX + 0.25, locY + 0.1875)
            shpObj.Ungroup()

            'Drop subtype title
            mastObj = stnObj.Masters("Subtypes for")
            shpObj = pagObj.Drop(mastObj, locX + 0.25, locY + 0.1875)
            shpObj.Text = "Subtypes of " & sClassName

            'Drop subtype field and default
            mastObj = stnObj.Masters("Subtype text")
            shpObj = pagObj.Drop(mastObj, locX + 1.0625, locY)
            shpObj.Text = sSubtypeField

            mastObj = stnObj.Masters("Subtype text")
            shpObj = pagObj.Drop(mastObj, locX + 1.0625, locY - 0.125)
            shpObj.Text = sDefaultSubtype

            'Drop subtype background
            mastObj = stnObj.Masters("Subtype background")
            shpObj = pagObj.Drop(mastObj, locX + 1.9375, locY - 0.5)
            shpObj.Cells("Height").ResultIU = numSubtypes * numUniqueValues * 0.125

            'Now, figure out the index of the subtype field, drop a blue outline around it, and make a line to the subtype section
            'MsgBox "Subtype field is " & sSubtypeField & ", index is " & lSubtypeFromBottom

            mastObj = stnObj.Masters("Subtype connector")
            shpObj = pagObj.Drop(mastObj, locX - 0.125, locY + 0.3125 + (lSubtypeFromBottom * 0.125))
            shpObj.Cells("Height").ResultIU = 0.2188 + (lSubtypeFromBottom * 0.125)

            'And, drop a blue box around the subtype field
            mastObj = stnObj.Masters("Subtype field")
            shpObj = pagObj.Drop(mastObj, locX, locY + 0.375 + (lSubtypeFromBottom * 0.125))

            locY = locY - 0.5

endOfFunction:

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Public Sub diagramSubtype(ByVal cCode As String, ByVal cName As String, ByVal numValues As Long, ByVal cFieldNames As String, ByVal cDefaultValues As String, ByVal cDomains As String)
        ' note: was a function

        Dim slashPos As Long
        Dim strLen As Long
        Dim i As Long

        Dim curFieldNames As String
        Dim curDefaultValues As String
        Dim curDomains As String
        Dim cThisFieldName As String
        Dim cThisDefaultValue As String
        Dim cThisDomain As String

        Dim boxHeight As Double

        Dim pWidthCode As Double
        Dim pWidthName As Double
        Dim pWidthSpacer As Double
        Dim pWidthFieldNames As Double
        Dim pWidthDefaultValues As Double
        Dim pWidthDomains As Double
        Dim pWidthTotal As Double

        Dim xCode As Double
        Dim xName As Double
        Dim xSpacer As Double
        Dim xFieldNames As Double
        Dim xDefaultValues As Double
        Dim xDomains As Double
        Dim thisY As Double
        Dim lNoValues As Boolean

        Try
            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            If dSets.Summary Then 'FormGDBDiagrammer.chkSummary.Value = 1 Then
                If lFirstSubtype = True Then
                    cNames = " are " & cName
                    lFirstSubtype = False
                    curY = curY - 0.3125
                Else
                    cNames = cNames & ", " & cName
                End If
                GoTo endOfFunction
            End If

            pWidthCode = 0.375
            pWidthName = 1.3125
            pWidthSpacer = 0.125
            pWidthFieldNames = 1.125
            pWidthDefaultValues = 1
            pWidthDomains = 1
            pWidthTotal = pWidthCode + pWidthName + pWidthSpacer + pWidthFieldNames + pWidthDefaultValues + pWidthDomains

            xCode = curX + 0.25
            xName = xCode + pWidthCode
            xSpacer = xName + pWidthName
            xFieldNames = xSpacer + pWidthSpacer
            xDefaultValues = xFieldNames + pWidthFieldNames
            xDomains = xDefaultValues + pWidthDefaultValues

            'Drop subtype and code, size boxes by number of values

            'For the case of no default/domain values, set equal to one
            lNoValues = False
            If numValues = 0 Or cFieldNames = "" Then
                numValues = 1
                lNoValues = True
            End If

            boxHeight = numValues * 0.125
            mastObj = stnObj.Masters("Field box")
            shpObj = pagObj.Drop(mastObj, xCode, locY)
            shpObj.Cells("Height").ResultIU = boxHeight
            shpObj.Cells("Width").ResultIU = pWidthCode
            shpObj.Text = cCode

            mastObj = stnObj.Masters("Field box")
            shpObj = pagObj.Drop(mastObj, xName, locY)
            shpObj.Cells("Height").ResultIU = boxHeight
            shpObj.Cells("Width").ResultIU = pWidthName
            shpObj.Text = cName

            curFieldNames = cFieldNames
            curDefaultValues = cDefaultValues
            curDomains = cDomains

            'If no values, then drop two field boxes and fill with note
            If lNoValues Then
                thisY = locY
                mastObj = stnObj.Masters("Field box")
                shpObj = pagObj.Drop(mastObj, xFieldNames, thisY)
                shpObj.Cells("Width").ResultIU = pWidthFieldNames
                shpObj.Text = "No values set"

                shpObj = pagObj.Drop(mastObj, xDefaultValues, thisY)
                shpObj.Cells("Width").ResultIU = pWidthDefaultValues
                shpObj.Text = ""

                shpObj = pagObj.Drop(mastObj, xDomains, thisY)
                shpObj.Cells("Width").ResultIU = pWidthDomains
                shpObj.Text = ""

            Else

                For i = 1 To numValues

                    thisY = locY - ((i - 1) * 0.125)
                    'Extract current field name
                    slashPos = InStr(1, curFieldNames, ":", 1)
                    strLen = Len(curFieldNames)
                    cThisFieldName = Left(curFieldNames, slashPos - 1)
                    curFieldNames = Right(curFieldNames, strLen - slashPos)

                    mastObj = stnObj.Masters("Field box")
                    shpObj = pagObj.Drop(mastObj, xFieldNames, thisY)
                    shpObj.Cells("Width").ResultIU = pWidthFieldNames
                    shpObj.Text = cThisFieldName

                    'Extract current default value
                    slashPos = InStr(1, curDefaultValues, ":", 1)
                    strLen = Len(curDefaultValues)
                    cThisDefaultValue = Left(curDefaultValues, slashPos - 1)
                    curDefaultValues = Right(curDefaultValues, strLen - slashPos)

                    mastObj = stnObj.Masters("Field box")
                    shpObj = pagObj.Drop(mastObj, xDefaultValues, thisY)
                    shpObj.Cells("Width").ResultIU = pWidthDefaultValues
                    shpObj.Text = cThisDefaultValue

                    'Extract current domain
                    slashPos = InStr(1, curDomains, ":", 1)
                    strLen = Len(curDomains)
                    cThisDomain = Left(curDomains, slashPos - 1)
                    curDomains = Right(curDomains, strLen - slashPos)

                    mastObj = stnObj.Masters("Field box")
                    shpObj = pagObj.Drop(mastObj, xDomains, thisY)
                    shpObj.Cells("Width").ResultIU = pWidthDomains
                    shpObj.Text = cThisDomain
                Next i
            End If

            'Drop outline box
            mastObj = stnObj.Masters("Subtype outline")
            shpObj = pagObj.Drop(mastObj, xCode, locY)
            shpObj.Cells("Width").ResultIU = pWidthTotal
            shpObj.Cells("Height").ResultIU = boxHeight

            'Drop arrow
            mastObj = stnObj.Masters("Subtype arrow")
            shpObj = pagObj.Drop(mastObj, xSpacer + (pWidthSpacer / 2), locY - (boxHeight / 2))

            locY = locY - (numValues * 0.125)
            curY = locY - 0.5

endOfFunction:

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Public Sub finishSubtype()
        'note:was a function
        Try

            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            Dim memberHeight As Double
            If dSets.Summary Then
                shpObj.Text = shpObj.Text & cNames
                memberHeight = shpObj.Cells("Height").ResultIU
                curY = curY - memberHeight + 0.3125
            End If
            lFirstSubtype = False

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Public Function startDomain(ByVal cDomainName As String, ByVal cDomainDescription As String, ByVal iFieldType As Integer,
                                                    ByVal iDomainType As Integer, ByVal iMergePolicy As Integer, ByVal iSplitPolicy As Integer, ByVal numDomainValues As Long, ByRef pDomain As IDomain)

        Dim pCodedValueDomain As ICodedValueDomain
        Dim pRangeDomain As IRangeDomain
        Dim pIndexDM As Long
        Dim cFieldType As String
#Disable Warning BC42024 ' Unused local variable: 'cDomainType'.
        Dim cDomainType As String
#Enable Warning BC42024 ' Unused local variable: 'cDomainType'.
        Dim cMergePolicy As String
        Dim cSplitPolicy As String

        Try


            Call positionOnPage()

            locX = curX
            locY = curY

            'Calculate and set height of box
            boxHeight = 0.8125 + (numDomainValues * 0.125)

            'Drop background
            mastObj = stnObj.Masters("Class background")
            shpObj = pagObj.Drop(mastObj, locX + 0.0313, locY - 0.0313)
            shpObj.Cells("Height").ResultIU = boxHeight
            shpObj.Cells("Width").ResultIU = 2.5

            'Drop domain class box
            mastObj = stnObj.Masters("Domain class box")
            shpObj = pagObj.Drop(mastObj, locX, locY)
            shpObj.Cells("Height").ResultIU = boxHeight
            If iDomainType = esriDomainType.esriDTCodedValue Then
                shpObj.Text = "Coded value domain"
            ElseIf iDomainType = esriDomainType.esriDTRange Then
                shpObj.Text = "Range domain"
            ElseIf iDomainType = esriDomainType.esriDTString Then
                shpObj.Text = "String domain"
            End If

            'Drop class name
            mastObj = stnObj.Masters("Class name")
            shpObj = pagObj.Drop(mastObj, locX, locY - 0.1403)
            shpObj.Text = cDomainName
            shpObj.Cells("Width").ResultIU = 2.5

            'Drop domain description and properties
            mastObj = stnObj.Masters("Domain description")
            shpObj = pagObj.Drop(mastObj, locX, locY - 0.2813)

            mastObj = stnObj.Masters("Domain properties")
            shpObj = pagObj.Drop(mastObj, locX + 0.5625, locY - 0.2813)


            cFieldType = findFieldType(iFieldType)

            Select Case iMergePolicy
                Case esriMergePolicyType.esriMPTSumValues
                    cMergePolicy = "Sum values"
                Case esriMergePolicyType.esriMPTAreaWeighted
                    cMergePolicy = "Area weighted"
                Case esriMergePolicyType.esriMPTDefaultValue
                    cMergePolicy = "Default value"
            End Select

            Select Case iSplitPolicy
                Case esriSplitPolicyType.esriSPTGeometryRatio
                    cSplitPolicy = "Geometry ratio"
                Case esriSplitPolicyType.esriSPTDuplicate
                    cSplitPolicy = "Duplicate"
                Case esriSplitPolicyType.esriSPTDefaultValue
                    cSplitPolicy = "Default value"
            End Select

#Disable Warning BC42104 ' Variable 'cMergePolicy' is used before it has been assigned a value. A null reference exception could result at runtime.
#Disable Warning BC42104 ' Variable 'cSplitPolicy' is used before it has been assigned a value. A null reference exception could result at runtime.
            shpObj.Text = cDomainDescription & vbCrLf & cFieldType & vbCrLf & cSplitPolicy & vbCrLf & cMergePolicy
#Enable Warning BC42104 ' Variable 'cSplitPolicy' is used before it has been assigned a value. A null reference exception could result at runtime.
#Enable Warning BC42104 ' Variable 'cMergePolicy' is used before it has been assigned a value. A null reference exception could result at runtime.

            Select Case iDomainType
                Case esriDomainType.esriDTCodedValue

                    'Drop field header
                    mastObj = stnObj.Masters("Domain header")
                    shpObj = pagObj.Drop(mastObj, locX, locY - 0.6875)

                    locY = curY - 0.8125

                    '----------------------------------------------------
                    ' Display Code Value Domain Values and Descriptions.
                    '----------------------------------------------------
                    pCodedValueDomain = pDomain
                    '
                    For pIndexDM = 0 To pCodedValueDomain.CodeCount - 1 Step 1
                        mastObj = stnObj.Masters("Domain box")
                        shpObj = pagObj.Drop(mastObj, curX, locY)
                        shpObj.Text = pCodedValueDomain.Value(pIndexDM)

                        shpObj = pagObj.Drop(mastObj, curX + 1.25, locY)
                        shpObj.Text = pCodedValueDomain.Name(pIndexDM)

                        locY = locY - 0.125
                    Next
                Case esriDomainType.esriDTRange
                    '----------------------------------------------
                    ' Display Range Domain Minimun/Maximum Values.
                    '----------------------------------------------
                    pRangeDomain = pDomain
                    mastObj = stnObj.Masters("Range header")
                    shpObj = pagObj.Drop(mastObj, locX, locY - 0.6875)
                    locY = curY - 0.8125

                    mastObj = stnObj.Masters("Domain box")
                    shpObj = pagObj.Drop(mastObj, curX, locY)
                    shpObj.Text = pRangeDomain.MinValue

                    shpObj = pagObj.Drop(mastObj, curX + 1.25, locY)
                    shpObj.Text = pRangeDomain.MaxValue

            End Select

            'drop outline box
            mastObj = stnObj.Masters("Box frame")
            shpObj = pagObj.Drop(mastObj, curX, curY)
            shpObj.Cells("Height").ResultIU = boxHeight
            shpObj.Cells("Width").ResultIU = 2.5

            curY = locY - 0.5

        Catch ex As Exception
            ExHandle(ex)
        End Try

#Disable Warning BC42105 ' Function 'startDomain' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'startDomain' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Function DiagramRelationship(ByVal cRelationshipClass As String, ByVal lIsManyToMany As Boolean, ByRef cRelProp1 As String, ByRef cRelProp2 As String, ByRef cRelProp3 As String,
                                                         ByRef cRelProp4 As String, ByRef iOriginType As Integer, ByRef iDestinationType As Integer, ByRef lAttributed As Boolean,
                                                         Optional ByRef pRelationshipTableFields As IFields = Nothing)

        Dim xRel1 As Double
        Dim xRel2 As Double
        Try


            Call positionOnPage()

            locX = curX
            locY = curY

            Dim dSets As DiagrammerSettings = FormGDBDiagrammer.pDiagrammerSettings

            'If summary checked, write out summary table shape...
            If dSets.Summary Then
                If lAttributed = True And lIsManyToMany = True Then
                    mastObj = stnObj.Masters("SmallAttRelMM")
                ElseIf lAttributed = True And lIsManyToMany = False Then
                    mastObj = stnObj.Masters("SmallAttRel1M")
                ElseIf lAttributed = False And lIsManyToMany = True Then
                    mastObj = stnObj.Masters("SmallRelMM")
                ElseIf lAttributed = False And lIsManyToMany = False Then
                    mastObj = stnObj.Masters("SmallRel1M")
                End If

                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Ungroup()
                mastObj = stnObj.Masters("RelationshipClassName")
                shpObj = pagObj.Drop(mastObj, locX + 0.2875, locY - 0.1563)
                shpObj.Text = cRelationshipClass

                locY = locY - 0.5
                curY = curY - 0.875
                GoTo endOfFunction
            End If

            'Calculate and set height of box
            boxHeight = 1.3125
            If lAttributed Then
                boxHeight = boxHeight + 0.25 + (pRelationshipTableFields.FieldCount * 0.125)
            End If

            'Drop background
            mastObj = stnObj.Masters("Class background")
            shpObj = pagObj.Drop(mastObj, locX + 0.0313, locY - 0.0313)
            shpObj.Cells("Height").ResultIU = boxHeight
            If lAttributed Then
                shpObj.Cells("Width").ResultIU = 4.8125
            Else
                shpObj.Cells("Width").ResultIU = 2.375
            End If

            'Drop relationship class box
            If lAttributed Then
                mastObj = stnObj.Masters("Relationship wide")
                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Cells("Height").ResultIU = boxHeight
                mastObj = stnObj.Masters("Relationship wide notes")
                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Ungroup()

            Else
                mastObj = stnObj.Masters("Relationship box")
                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Cells("Height").ResultIU = boxHeight
                mastObj = stnObj.Masters("Relationship notes")
                shpObj = pagObj.Drop(mastObj, locX, locY)
                shpObj.Ungroup()

            End If

            'Drop class name
            mastObj = stnObj.Masters("Class name")
            shpObj = pagObj.Drop(mastObj, locX + 0.2813, locY - 0.1563)
            shpObj.Text = cRelationshipClass

            mastObj = stnObj.Masters("Relationship origin")
            shpObj = pagObj.Drop(mastObj, locX, locY - 0.6563)
            If iOriginType = esriDatasetType.esriDTFeatureClass Then
                shpObj.Text = "Origin feature class"
            Else
                shpObj.Text = "Origin table"
            End If
            If lAttributed Then shpObj.Cells("Width").ResultIU = 2.3438

            mastObj = stnObj.Masters("Relationship destination")
            If lAttributed Then
                shpObj = pagObj.Drop(mastObj, locX + 2.4688, locY - 0.6563)
                shpObj.Cells("Width").ResultIU = 2.3428
            Else
                shpObj = pagObj.Drop(mastObj, locX + 1.25, locY - 0.6563)
            End If

            If iDestinationType = esriDatasetType.esriDTFeatureClass Then
                shpObj.Text = "Destination feature class"
            Else
                shpObj.Text = "Destination table"
            End If

            If lAttributed Then
                xRel1 = 0.5938
                xRel2 = 3.0938
            Else
                xRel1 = 0.4688
                xRel2 = 1.7031
            End If

            mastObj = stnObj.Masters("Relationship properties")
            shpObj = pagObj.Drop(mastObj, locX + xRel1, locY - 0.2813)
            shpObj.Text = cRelProp1

            shpObj = pagObj.Drop(mastObj, locX + xRel2, locY - 0.2813)
            shpObj.Text = cRelProp2

            shpObj = pagObj.Drop(mastObj, locX + xRel1, locY - 0.7813)
            shpObj.Text = cRelProp3

            shpObj = pagObj.Drop(mastObj, locX + xRel2, locY - 0.7813)
            shpObj.Text = cRelProp4

            mastObj = stnObj.Masters("Relationship description")
            If lAttributed Then
                shpObj = pagObj.Drop(mastObj, locX + 2.6172, locY - 0.7813)
            Else
                shpObj = pagObj.Drop(mastObj, locX + 1.25, locY - 0.7813)
            End If
            If lIsManyToMany Then
                shpObj.Text = "Name" & vbCrLf & "Primary key" & vbCrLf & "Foreign key"
            Else
                shpObj.Text = "Name"
            End If

            'Drop relationship connectors
            mastObj = stnObj.Masters("Relationship connector")
            shpObj = pagObj.Drop(mastObj, locX - 0.25, locY - 0.4375)
            If lAttributed Then
                shpObj = pagObj.Drop(mastObj, locX + 5.0625, locY - 1.0#)
            Else
                shpObj = pagObj.Drop(mastObj, locX + 2.625, locY - 1.0#)
            End If

            If lAttributed Then
                'Drop field header
                mastObj = stnObj.Masters("Relationship header")
                shpObj = pagObj.Drop(mastObj, locX, locY - 1.3125)
                shpObj.Ungroup()
                locY = locY - 1
                Call DiagramFeatureClass(pRelationshipTableFields, cRelationshipClass)
                curY = curY - 0.25
            Else
                'DiagramFeatureClass drops an outline box, so we must add one manually here...
                mastObj = stnObj.Masters("Box frame")
                shpObj = pagObj.Drop(mastObj, curX, locY)
                boxHeight = 1.3125
                shpObj.Cells("Height").ResultIU = boxHeight
                shpObj.Cells("Width").ResultIU = 2.375
                curY = curY - boxHeight - 0.5
            End If

endOfFunction:

        Catch ex As Exception
            ExHandle(ex)
        End Try

#Disable Warning BC42105 ' Function 'DiagramRelationship' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.
    End Function
#Enable Warning BC42105 ' Function 'DiagramRelationship' doesn't return a value on all code paths. A null reference exception could occur at run time when the result is used.

    Public Function setFieldType(ByVal cTypeX As String, ByVal cFieldName As String, ByVal lValid As Boolean) As String

        Dim result As String = ""
        Try

            Dim cUCaseFieldName As String = UCase(cFieldName)

            'Set normal field box to be the default shape type
            result = "Field box"

            'If field name is object id, shape, shape length, or shape area, set to reserved field type
            If cUCaseFieldName = "OBJECTID" Or cUCaseFieldName = "SHAPE" Then
                result = "Reserved field box"
            End If

            If lValid = False Then
                result = "Null field box"
            End If


            Select Case curClassType

                Case esriFeatureType.esriFTSimple
                    If cUCaseFieldName = "SHAPE_LENGTH" Or cUCaseFieldName = "SHAPE_AREA" Then result = "Reserved field box"

                Case esriFeatureType.esriFTComplexEdge
                    If cUCaseFieldName = "SHAPE_LENGTH" Or cUCaseFieldName = "ENABLED" Then result = "Reserved field box"

                Case esriFeatureType.esriFTSimpleJunction
                    If cUCaseFieldName = "ENABLED" Then result = "Reserved field box"

                Case esriFeatureType.esriFTSimpleEdge
                    If cUCaseFieldName = "SHAPE_LENGTH" Or cUCaseFieldName = "ENABLED" Then result = "Reserved field box"

                Case esriFeatureType.esriFTAnnotation
                    If cUCaseFieldName = "SHAPE_LENGTH" Or cUCaseFieldName = "SHAPE_AREA" Or
                        cUCaseFieldName = "FEATUREID" Or cUCaseFieldName = "ZORDER" Or cUCaseFieldName = "ANNOTATIONCLASSID" Or
                        cUCaseFieldName = "ELEMENT" Then result = "Reserved field box"

                Case esriFeatureType.esriFTDimension
                    If cUCaseFieldName = "SHAPE_LENGTH" Or cUCaseFieldName = "SHAPE_AREA" Or
                        cUCaseFieldName = "DIMLENGTH" Or cUCaseFieldName = "BEGINX" Or cUCaseFieldName = "BEGINY" Or
                        cUCaseFieldName = "ENDX" Or cUCaseFieldName = "ENDY" Or cUCaseFieldName = "DIMX" Or
                        cUCaseFieldName = "DIMY" Or cUCaseFieldName = "TEXTX" Or cUCaseFieldName = "TEXTY" Or
                        cUCaseFieldName = "DIMTYPE" Or cUCaseFieldName = "EXTANGLE" Or cUCaseFieldName = "STYLEID" Or
                        cUCaseFieldName = "USECUSTOMLENGTH" Or cUCaseFieldName = "CUSTOMLENGTH" Or cUCaseFieldName = "DIMDISPLAY" Or
                        cUCaseFieldName = "EXTDISPLAY" Or cUCaseFieldName = "MARKERDISPLAY" Or cUCaseFieldName = "TEXTANGLE" Then result = "Reserved field box"

            End Select

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result

    End Function

    Public Function findFieldType(ByVal iFieldType As Integer) As String

        Dim result As String = ""

        Try
            Select Case iFieldType
                Case esriFieldType.esriFieldTypeSmallInteger
                    result = "Short integer"
                Case esriFieldType.esriFieldTypeInteger
                    result = "Long integer"
                Case esriFieldType.esriFieldTypeSingle
                    result = "Float"
                Case esriFieldType.esriFieldTypeDouble
                    result = "Double"
                Case esriFieldType.esriFieldTypeString
                    result = "String"
                Case esriFieldType.esriFieldTypeOID
                    result = "Object ID"
                Case esriFieldType.esriFieldTypeGeometry
                    result = "Geometry"
                Case esriFieldType.esriFieldTypeDate
                    result = "Date"
                Case esriFieldType.esriFieldTypeBlob
                    result = "Blob"
            End Select

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result

    End Function

    Public Sub positionOnPage()
        ' was a function

        Try

            Dim vertSep As Double, columnWidth As Double
            vertSep = 0.125
            columnWidth = 8

            'If curX = 0, then initialize
            If curX = 0 Then
                curX = 1
                curY = 33
            End If

            'Check if the new height will take us beyond the bottom of the drawable area.
            'If so, start a new column
            If curY < 5 Then
                curY = 35
                curX = curX + columnWidth
            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Public Sub CloseDiagram()
        Try

            docObj.SaveAs(cVisioFileName)

            docObj.Close()
            appVisio.Quit()

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

#End Region

End Module
