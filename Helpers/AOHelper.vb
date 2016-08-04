Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Output
Imports ESRI.ArcGIS.Display


'---------------------
' Geodatabase RDBMS's
'---------------------
Public Enum enumGeodatabaseFlavor
    enumGFAccess = 0
    enumGFOracle = 1
    enumGFSQLServer = 2
    enumGFInformix = 3
    enumGFDB2 = 4

    enumGFUnknown = -1
    enumGFInvalid = -2
End Enum
Structure structFixedLengthString
    ' This is for the fixed length string used in GetSpatialReferenceDescription
    <VBFixedString(2048)> Friend pFixBuffer As String

End Structure

Module AOHelper


#Region "ArcObjects"

    Public Function GetRasterDatasetFromWorkspace(ByRef pWorkspace As IWorkspace, _
                                              ByRef pDatasetName As String) As IRasterDataset
        Dim result As IRasterDataset = Nothing
        Try

            '-------------------------------------------------------------------
            ' Returns an IRasterDataset interface from an IWorkspace interface.
            ' The other alternative is to make a RasterWorkspaceFactory object!
            '-------------------------------------------------------------------
            Dim pEnumDataset As IEnumDataset
            Dim pDataset As IDataset

            pEnumDataset = pWorkspace.Datasets(esriDatasetType.esriDTRasterDataset)
            pDataset = pEnumDataset.Next
            Do Until pDataset Is Nothing
                If UCase(pDataset.Name) = UCase(pDatasetName) Then
                    result = pDataset
                    Exit Do
                End If
                pDataset = pEnumDataset.Next
            Loop

        Catch ex As Exception
            ExHandle(ex)
        End Try
        Return result
    End Function

    Public Function GetDatasetFromName(ByRef pWorkspace As IWorkspace, _
                                       ByRef pDatasetName As String, _
                                       ByRef pDatasetType As esriDatasetType) As IDataset
        Dim result As IDataset = Nothing
        Try


            '--------------------------------------------------------------------
            ' Fetches the IDataset interface based on the dataset name and type.
            '--------------------------------------------------------------------
            Dim pFeatureWorkspace As IFeatureWorkspace
            '
            Select Case pDatasetType
                Case esriDatasetType.esriDTTerrain
                    ' Not YET IMPLEMENTED
                Case esriDatasetType.esriDTGeometricNetwork
                    ' Not YET IMPLEMENTED


                Case esriDatasetType.esriDTFeatureClass
                    pFeatureWorkspace = pWorkspace
                    result = pFeatureWorkspace.OpenFeatureClass(pDatasetName)

                Case esriDatasetType.esriDTFeatureDataset
                    pFeatureWorkspace = pWorkspace
                    result = pFeatureWorkspace.OpenFeatureDataset(pDatasetName)
                Case esriDatasetType.esriDTRasterDataset
                    '--------------------------------------------------------------------------------
                    ' Call routine to search through each standalone dataset until the name matches.
                    '--------------------------------------------------------------------------------
                    result = GetRasterDatasetFromWorkspace(pWorkspace, pDatasetName)
                Case esriDatasetType.esriDTTable
                    pFeatureWorkspace = pWorkspace
                    result = pFeatureWorkspace.OpenTable(pDatasetName)
                Case esriDatasetType.esriDTRelationshipClass
                    pFeatureWorkspace = pWorkspace
                    result = pFeatureWorkspace.OpenRelationshipClass(pDatasetName)

            End Select

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result

    End Function

    Public Function GetSpatialReferenceFromDataset(ByRef pDataset As IDataset) As ISpatialReference2

        Dim result As ISpatialReference2 = Nothing
        Try

            '-----------------------------------------------------------
            ' Extract the SpatialReference from the IDataset interface.
            '-----------------------------------------------------------
            Dim pGeoDataset As IGeoDataset = TryCast(pDataset, IGeoDataset)
            result = pGeoDataset.SpatialReference

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result

    End Function

    Public Function GetSpatialReferenceDescription(ByRef pSpatialReference As ISpatialReference2) As String

        Dim result As String = ""
        Try
            '---------------------------------------------------------------------------------------
            ' Returns the Spatial Reference description from the parsed spatial reference iterface.
            ' Also trims off the chr(0)'s from the buffered string.
            '---------------------------------------------------------------------------------------

            ' Change this:
            ' Dim pBuffer As String * 2048
            ' To this:
            Dim pFixedLengthString As New structFixedLengthString
            pFixedLengthString.pFixBuffer = ""
            Dim pBuffer = pFixedLengthString.pFixBuffer

            Dim pBytes As Long = Nothing

            Dim pESRISpatialReference As IESRISpatialReferenceGEN = CType(pSpatialReference, ISpatialReference2GEN)

            pESRISpatialReference.ExportToESRISpatialReference(pBuffer, pBytes)

            result = Left(pBuffer, InStr(1, pBuffer, Chr(0)) - 1)

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result
    End Function

    Public Function ReformatProjectionString(ByRef pProjectionString As String) As String

        Dim result As String = ""
        Try

            '--------------------------------------------------------------------
            ' Formats the OGC Well-Known Text Representation of Spatial Systems.
            ' Adds "Carriage Returns" and "Spaces".
            '--------------------------------------------------------------------
            Dim pProjectionSplit As Object
            Dim pProjectionStringOut As String
            Dim pIndex As Long
            Dim pProjectionStringPart As String
#Disable Warning BC42024 ' Unused local variable: 'pTabLevel'.
            Dim pTabLevel As Integer
#Enable Warning BC42024 ' Unused local variable: 'pTabLevel'.
            '-----------------------------------------------------------
            ' Create a list of nodes (separated by a <br> - breakline).
            '-----------------------------------------------------------
            pProjectionSplit = Split(pProjectionString, ",")
            pProjectionStringOut = ""
            '
            For pIndex = LBound(pProjectionSplit, 1) To UBound(pProjectionSplit, 1) Step 1
                pProjectionStringPart = CStr(pProjectionSplit(pIndex))
                If pProjectionStringPart <> "" Then
                    If InStr(1, pProjectionStringPart, "[") <> 0 Then
                        If pProjectionStringOut = "" Then
                            pProjectionStringOut = pProjectionStringPart
                        Else
                            pProjectionStringOut = pProjectionStringOut & "<br>" & pProjectionStringPart
                        End If
                    Else
                        If pProjectionStringOut = "" Then
                            pProjectionStringOut = pProjectionStringPart
                        Else
                            pProjectionStringOut = pProjectionStringOut & "," & pProjectionStringPart
                        End If
                    End If
                End If
            Next pIndex

            result = pProjectionStringOut

        Catch ex As Exception
            ExHandle(ex)
        End Try
        Return result
    End Function

    Public Sub MakePictureOfDataset(ByRef pDataset As IDataset, _
                                    ByRef pSizeX As Long, _
                                    ByRef pSizeY As Long, _
                                    ByRef pResolution As Long, _
                                    ByRef pOutFilename As String, _
                                    Optional ByRef pBackgroundColor As IColor = Nothing)
        Try

            Dim pLayer As ILayer
            Dim pMap As IMap
            Dim pActiveView As IActiveView
            Dim pExporter As IExporter
            Dim pEnvelope As IEnvelope

            Dim pJpegExporter As IJpegExporter
            Dim pFeatureLayer As IFeatureLayer2
            Dim pDimensionLayer As IDimensionLayer
            Dim pGeoFeatureLayer As IGeoFeatureLayer
            Dim pRasterLayer As IRasterLayer
            Dim pFeatureClass As IFeatureClass

            Select Case pDataset.Type

                Case esriDatasetType.esriDTFeatureClass
                    pFeatureClass = pDataset

                    Select Case pFeatureClass.FeatureType
                        Case esriFeatureType.esriFTDimension
                            pDimensionLayer = New DimensionLayer
                            pFeatureLayer = pDimensionLayer

                        Case esriFeatureType.esriFTAnnotation
                            pFeatureLayer = New FeatureLayer
                            pGeoFeatureLayer = pFeatureLayer
                            pGeoFeatureLayer.DisplayAnnotation = True

                        Case Else
                            pFeatureLayer = New FeatureLayer
                    End Select

                    pFeatureLayer.FeatureClass = pFeatureClass
                    pLayer = pFeatureLayer
                Case esriDatasetType.esriDTRasterDataset
                    pRasterLayer = New RasterLayerClass
                    pRasterLayer.CreateFromDataset(pDataset)
                    pLayer = pRasterLayer
                Case Else
                    MsgBox("Cannot create a thumbnail for dataset: [" & pDataset.Name & "]", vbCritical) ', App.ProductName)
                    Exit Sub
            End Select

            pMap = New Map
            pMap.AddLayer(pLayer)

            pActiveView = pMap
            pExporter = New JpegExporter
            pEnvelope = New Envelope

            pExporter.Resolution = pResolution
            pExporter.ExportFileName = pOutFilename

            pJpegExporter = pExporter
            If Not pBackgroundColor Is Nothing Then
                pJpegExporter.BackgroundColor = pBackgroundColor
            End If
            pJpegExporter.Height = pSizeY
            pJpegExporter.Width = pSizeX
            pJpegExporter.Quality = 100

            Dim pTagRect As New ESRI.ArcGIS.esriSystem.tagRECT
            pTagRECT.Left = 0
            pTagRECT.Top = 0
            pTagRECT.bottom = pSizeY
            pTagRECT.Right = pSizeX
            pActiveView.Output(pExporter.StartExporting, pResolution, pTagRECT, Nothing, Nothing)
            pExporter.FinishExporting()

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Public Function GetRowCount(ByRef pTable As ITable, _
                                Optional ByRef pSubtypeCode As Long = -999999999) As Long
        Dim result As Long = 0
        Try

            '----------------------------------------------------------------------------
            ' Returns a count of the rows (or features). Optional specific Subtype Code.
            '----------------------------------------------------------------------------
            If pSubtypeCode = -999999999 Then
                GetRowCount = pTable.RowCount(Nothing)
            Else
                Dim pSubtypes As ISubtypes
                pSubtypes = pTable

                Dim pQueryFilter As IQueryFilter2
                pQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = pSubtypes.SubtypeFieldName & " = " & pSubtypeCode
                GetRowCount = pTable.RowCount(pQueryFilter)
            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try
        Return result

    End Function

    Public Function GetGeodatabaseFlavor(ByRef pWorkspace As IWorkspace) As enumGeodatabaseFlavor

        Dim result As enumGeodatabaseFlavor = enumGeodatabaseFlavor.enumGFUnknown
        Try

            '---------------------------------------------------
            ' This function will return the Geodatabase's RDBMS
            '---------------------------------------------------
            Const pKeywordOracle As String = "ACCESS"
            Const pKeywordSQLServer As String = "AUTHORIZATION"
            Const pKeywordInformix As String = "?"
            Const pKeywordDB2 As String = "?"
            '
            Select Case pWorkspace.Type

                Case esriWorkspaceType.esriFileSystemWorkspace
                    result = enumGeodatabaseFlavor.enumGFInvalid
                Case esriWorkspaceType.esriLocalDatabaseWorkspace
                    result = enumGeodatabaseFlavor.enumGFAccess
                Case esriWorkspaceType.esriRemoteDatabaseWorkspace
                    '--------------------------------------
                    ' Select RDBMS based on reserved words
                    '--------------------------------------
                    Dim pSQLSyntax As ISQLSyntax
                    Dim pEnumBSTR As IEnumBSTR
                    Dim pKeyword As String
#Disable Warning BC42024 ' Unused local variable: 'pTable'.
                    Dim pTable As ITable
#Enable Warning BC42024 ' Unused local variable: 'pTable'.
                    '
                    pSQLSyntax = pWorkspace
                    pEnumBSTR = pSQLSyntax.GetKeywords
                    pKeyword = pEnumBSTR.Next
                    result = enumGeodatabaseFlavor.enumGFUnknown

                    Do Until pKeyword = ""
                        Select Case pKeyword
                            Case pKeywordOracle
                                result = enumGeodatabaseFlavor.enumGFOracle
                            Case pKeywordSQLServer
                                result = enumGeodatabaseFlavor.enumGFSQLServer
                            Case pKeywordInformix
                                result = enumGeodatabaseFlavor.enumGFInformix
                            Case pKeywordDB2
                                result = enumGeodatabaseFlavor.enumGFDB2
                        End Select
                        pKeyword = pEnumBSTR.Next
                    Loop
            End Select


        Catch ex As Exception
            ExHandle(ex)
        End Try
        Return result

    End Function


#End Region

#Region "VB Utils"

    Friend Sub ExHandle(ByVal ex As Exception)
        Try
            Dim sStack As String = ex.StackTrace
            MsgBox("Source:" & sStack & vbNewLine & "Msg: " & ex.Message, MsgBoxStyle.Information, ex.Source)

        Catch ex2 As Exception
            MsgBox("Error with the ExHandle exception handler.")
        End Try
    End Sub
#End Region

End Module
