Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcCatalog
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Catalog

Public Class FormGDBDiagrammer

#Region "Member Variables"

    '-------------------------------------------------------------------------------
    ' Reg Key Security Options. Used by RegOpenKeyEx, RegQueryValueEx & RegCloseKey
    '-------------------------------------------------------------------------------
    Private Const READ_CONTROL = &H20000
    Private Const KEY_QUERY_VALUE = &H1
    Private Const KEY_SET_VALUE = &H2
    Private Const KEY_CREATE_SUB_KEY = &H4
    Private Const KEY_ENUMERATE_SUB_KEYS = &H8
    Private Const KEY_NOTIFY = &H10
    Private Const KEY_CREATE_LINK = &H20
    Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                                   KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                                   KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

    '-------------------------------------------------------------------------
    ' Reg Key ROOT Types. Used by RegOpenKeyEx, RegQueryValueEx & RegCloseKey
    '-------------------------------------------------------------------------
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const ERROR_SUCCESS = 0
    Private Const REG_SZ = 1                  ' Unicode null terminated string
    Private Const REG_DWORD = 4               ' 32-bit number
    Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
    Private Const gREGVALSYSINFOLOC = "MSINFO"
    Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
    Private Const gREGVALSYSINFO = "PATH"

    '--------------------------------------------------------------------
    ' Constants User by SHGetSpecialFolderLocation & SHGetPathFromIDList
    '--------------------------------------------------------------------
    Private Const CSIDL_DESKTOP = &H0
    Private Const CSIDL_PROGRAMS = &H2
    Private Const CSIDL_CONTROLS = &H3
    Private Const CSIDL_PRINTERS = &H4
    Private Const CSIDL_PERSONAL = &H5
    Private Const CSIDL_FAVORITES = &H6
    Private Const CSIDL_STARTUP = &H7
    Private Const CSIDL_RECENT = &H8
    Private Const CSIDL_SENDTO = &H9
    Private Const CSIDL_BITBUCKET = &HA
    Private Const CSIDL_STARTMENU = &HB
    Private Const CSIDL_DESKTOPDIRECTORY = &H10
    Private Const CSIDL_DRIVES = &H11
    Private Const CSIDL_NETWORK = &H12
    Private Const CSIDL_NETHOOD = &H13
    Private Const CSIDL_FONTS = &H14
    Private Const CSIDL_TEMPLATES = &H15
    Private Const MAX_PATH = 260

    '----------------------------------------------------------------
    ' Types used by SHGetSpecialFolderLocation & SHGetPathFromIDList
    '----------------------------------------------------------------
    'Private Type SHITEMID
    '    cb As Long
    '    abID As Byte
    'End Type
    'Private Type ITEMIDLIST
    '    mkid As SHITEMID
    'End Type
    Structure SHITEMID
        Dim cb As Long
        Dim abID As Byte
    End Structure

    Private Structure ITEMIDLIST
        Dim mkid As SHITEMID
    End Structure

    '----------------------------------------------------------------------
    ' Constant used by ShellExecute to launch default windows applications
    '----------------------------------------------------------------------
    Private Const SW_SHOW = 5

    '-----------------------------------
    ' Windows API Function Declarations
    '-----------------------------------
    Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
    Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
    Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal pidl As ITEMIDLIST) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    '-----------------------------------------------------------------------------------
    ' Variables that hook into the ArcCatalog Application and the selected Geodatabase.
    '------------------------------------------------------------------------------------
    Private mGxApplication As IGxApplication
    Private mWorkspace As IWorkspace

    '-------------------------------------------------------------
    ' String Array that contains all the language cross-reference
    '-------------------------------------------------------------
    ' CHANGED: removed this because it wasn't being used anywhere.
    'Private mLanguageReference() As String

    '---------------------------------------------------------------
    ' Dimensions for Image Snapshots (Row/Feature/Cell Count Table)
    '---------------------------------------------------------------
    Private Const mSnapshotSmall As Long = 50
    Private Const mSnapshotBig As Long = 800
    Private Const mSnapshotResolution As Long = 150

    '---------------------------------------------------
    ' Use this string as a "No Data" flag rather and ""
    '---------------------------------------------------
    Private Const mNoData As String = "-999999999"

    '--------------------------------
    ' Form spacing between controls.
    '--------------------------------
    Private Const mFormGridSpacing As Integer = 120

#End Region

#Region "Public Members"

    'Friend Shared bSummary As Boolean = True

    'Public g_FormSettings As System.Collections.Specialized.StringDictionary
    'Public g_sPSorTT As String

    Friend Shared pDiagrammerSettings As DiagrammerSettings

#End Region

#Region "Event Handlers"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        pDiagrammerSettings = New DiagrammerSettings

        With pDiagrammerSettings
            .Postcript = True
            .TrueType = False
            .UseAbstract = False
            .FieldMetadata = False
            .FieldAlias = False
            .OmitAnno = False
            .Summary = True
        End With

        Me.optPS.Checked = True

        Me.btnRunGDBDiagrammer.Enabled = False

        ' the hook
        mGxApplication = My.ThisApplication

    End Sub

    Private Sub btnRunGDBDiagrammer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRunGDBDiagrammer.Click
        Try
            ' Generate the Diagram.

            Dim pEnumDataset As IEnumDataset
            Dim pEnumDataset2 As IEnumDataset
            Dim pDataset As IDataset
            Dim pDataset2 As IDataset
#Disable Warning BC42024 ' Unused local variable: 'pNetworkCollection'.
            Dim pNetworkCollection As INetworkCollection
#Enable Warning BC42024 ' Unused local variable: 'pNetworkCollection'.
#Disable Warning BC42024 ' Unused local variable: 'pGeometricNetwork'.
            Dim pGeometricNetwork As IGeometricNetwork
#Enable Warning BC42024 ' Unused local variable: 'pGeometricNetwork'.
#Disable Warning BC42024 ' Unused local variable: 'pCounter'.
            Dim pCounter As Long
#Enable Warning BC42024 ' Unused local variable: 'pCounter'.
            Dim pWorkspaceDomains As IWorkspaceDomains
            Dim pEnumDomain As IEnumDomain
            Dim pDomain As IDomain
            Dim pGxObject As IGxObject

#Disable Warning BC42024 ' Unused local variable: 'pIndex'.
            Dim pIndex As Long
#Enable Warning BC42024 ' Unused local variable: 'pIndex'.
            ' Dim pFolder As Scripting.Folder
#Disable Warning BC42024 ' Unused local variable: 'pFolderName'.
            Dim pFolderName As String
#Enable Warning BC42024 ' Unused local variable: 'pFolderName'.

#Disable Warning BC42024 ' Unused local variable: 'pGxApplication'.
            Dim pGxApplication As IGxApplication
#Enable Warning BC42024 ' Unused local variable: 'pGxApplication'.

            Me.ToolStripStatusLabel1.Text = "Starting ..."
            Me.ToolStripStatusLabel1.Invalidate()
            System.Windows.Forms.Application.DoEvents()



            'Set the hourglass cursor
            Dim pMouseCursor As ESRI.ArcGIS.Framework.IMouseCursor
            pMouseCursor = New ESRI.ArcGIS.Framework.MouseCursor
            pMouseCursor.SetCursor(2)


            '---------------------------------
            ' Reset the IWorkspace Interface.
            '---------------------------------
            pGxObject = mGxApplication.SelectedObject
            If TypeOf pGxObject Is IGxDatabase Then
                Dim pGxDatabase As IGxDatabase
                pGxDatabase = pGxObject
                mWorkspace = pGxDatabase.Workspace
            Else
                MsgBox("Please Select A Personal/ArcSDE Geodatabase", vbExclamation)
                Exit Sub
            End If
            '----------------------------
            ' Populate "Dataset" ListBox
            '----------------------------
            Me.lstDataset.Items.Clear()
            pEnumDataset = mWorkspace.Datasets(esriDatasetType.esriDTAny)
            pDataset = pEnumDataset.Next
            Do Until pDataset Is Nothing
                If pDataset.Type = esriDatasetType.esriDTFeatureDataset Then
                    pEnumDataset2 = pDataset.Subsets
                    pDataset2 = pEnumDataset2.Next
                    Do Until pDataset2 Is Nothing
                        '-----------------------------------------------------------------------
                        ' FD Entry:            0   |  FD Name  |  Dataset Name  |  DatasetType
                        '-----------------------------------------------------------------------
                        Me.lstDataset.Items.Add("0" & "|" & pDataset.Name & "|" & pDataset2.Name & "|" & pDataset2.Type)
                        pDataset2 = pEnumDataset2.Next
                    Loop
                Else
                    '--------------------------------------------------------------------
                    ' No FD Entry:         1   |        |  Dataset Name  |  DatasetType
                    '--------------------------------------------------------------------
                    Me.lstDataset.Items.Add("1" & "|" & "|" & pDataset.Name & "|" & pDataset.Type)
                End If
                pDataset = pEnumDataset.Next
            Loop
            Me.ToolStripStatusLabel1.Text = "Finished Setup Dataset"
            Me.ToolStripStatusLabel1.Invalidate()
            System.Windows.Forms.Application.DoEvents()


            '--------------------------
            ' Populate Domain ListBox.
            '--------------------------
            Me.lstDomain.Items.Clear()
            pWorkspaceDomains = mWorkspace
            pEnumDomain = pWorkspaceDomains.Domains
            If Not pEnumDomain Is Nothing Then
                pDomain = pEnumDomain.Next
                Do Until pDomain Is Nothing
                    Me.lstDomain.Items.Add(pDomain.Name & "|" & CStr(pDomain.Type))
                    pDomain = pEnumDomain.Next
                Loop
            End If

            Me.ToolStripStatusLabel1.Text = "Finished Setup Domains"
            Me.ToolStripStatusLabel1.Invalidate()
            System.Windows.Forms.Application.DoEvents()

            '-----------------------------------
            ' Start diagramming tool. MNZ
            '-----------------------------------
            Me.chkSummary.Checked = True
            pDiagrammerSettings.Summary = True

            Me.ToolStripStatusLabel1.Text = "Creating the Diagram ...."
            Me.ToolStripStatusLabel1.Invalidate()
            System.Windows.Forms.Application.DoEvents()


            VisioDiagram.StartDiagram(mWorkspace.PathName, Me.txtOutputFile.Text)

            VisioDiagram.SectionHeader()
            WriteObjectClassInformation()
            WriteRelationshipClassInformation()

            Me.chkSummary.Checked = False
            pDiagrammerSettings.Summary = False

            VisioDiagram.SectionHeader()
            WriteObjectClassInformation()
            WriteRelationshipClassInformation()
            WriteDomainInformation()

            '--------------------------
            ' Clear ListBoxes Controls
            '--------------------------
            Me.lstDataset.Items.Clear()
            Me.lstDomain.Items.Clear()
            Me.lstTemp.Items.Clear()

            '-----------------------------------
            ' Stop diagramming tool. MNZ
            '-----------------------------------
            VisioDiagram.CloseDiagram()

            Me.ToolStripStatusLabel1.Text = "Completed the diagram."
            System.Windows.Forms.Application.DoEvents()

            pMouseCursor = Nothing

            ClearModularVariables()

            '-----------------------------------
            ' Close the form and exit
            '-----------------------------------
            'Me.Hide()
            'Me.Dispose()
            ' Unload(Me)

            'MsgBox("The Geodatabse Diagrammer successfully completed.", MsgBoxStyle.Information, "GDB Diagrammer")

        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub txtOutputFile_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOutputFile.TextChanged
        Try
            If Me.txtOutputFile.Text <> "" Then
                Me.btnRunGDBDiagrammer.Enabled = True
            Else
                Me.btnRunGDBDiagrammer.Enabled = False
            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub btnSaveFileDialog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFileDialog.Click
        ' Open the Save File dialog.
        Try

            Dim sFileName As String = ""

            With SaveFileDialog1
                .FileName = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GeodatabaseDiagram.vsd")
                .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                .Filter = "Visio Diagram (*.vsd)|*.vsd"

                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    sFileName = .FileName
                    Me.txtOutputFile.Text = sFileName
                End If
            End With

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        Try
            Me.Hide()
            Me.Dispose()
        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Private Sub chkAbstract_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAbstract.CheckedChanged
        Try
            pDiagrammerSettings.UseAbstract = Me.chkAbstract.Checked
        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub


    Private Sub chkFieldMetadataD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFieldMetadataD.CheckedChanged
        Try
            pDiagrammerSettings.FieldMetadata = Me.chkFieldMetadataD.Checked
        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub chkFieldAlias_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFieldAlias.CheckedChanged
        Try
            pDiagrammerSettings.FieldAlias = Me.chkFieldAlias.Checked
        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub chkOmitAnno_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOmitAnno.CheckedChanged
        Try
            pDiagrammerSettings.OmitAnno = Me.chkOmitAnno.Checked
        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub optTT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTT.CheckedChanged
        Try
            pDiagrammerSettings.TrueType = optTT.Checked
        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub optPS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPS.CheckedChanged
        Try
            pDiagrammerSettings.Postcript = optPS.Checked
        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

#End Region

#Region "Private Functions"

    Private Function GetKeyValue(ByVal KeyRoot As Long, ByVal KeyName As String, ByVal SubKeyRef As String, ByRef KeyVal As String) As Boolean

        Dim result As Boolean = False

        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
#Disable Warning BC42024 ' Unused local variable: 'hDepth'.
        Dim hDepth As Long                                      '
#Enable Warning BC42024 ' Unused local variable: 'hDepth'.
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String = ""                                 ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable

        Try
            '------------------------------------------------------------
            ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
            '------------------------------------------------------------
            rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
            'If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Handle Error...
            If (rc <> ERROR_SUCCESS) Then
                Dim up As New Exception("Error Opening the Registry Key")
                Throw up
            End If

            'tmpVal = String$(1024, 0)                     ' Allocate Variable Space
            tmpVal.PadLeft(1024, " "c)
            KeyValSize = 1024                              ' Mark Variable Size

            '------------------------------------------------------------
            ' Retrieve Registry Key Value...
            '------------------------------------------------------------
            rc = RegQueryValueEx(hKey, SubKeyRef, 0,
                                 KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

            'If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Handle Errors
            If (rc <> ERROR_SUCCESS) Then
                Dim up As New Exception("Error Getting the Registry Key value.")
                Throw up
            End If



            If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
                tmpVal = Strings.Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
            Else                                                    ' WinNT Does NOT Null Terminate String...
                tmpVal = Strings.Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
            End If


            '------------------------------------------------------------
            ' Determine Key Value Type For Conversion...
            '------------------------------------------------------------
            Select Case KeyValType                                  ' Search Data Types...
                Case REG_SZ                                             ' String Registry Key Data Type
                    KeyVal = tmpVal                                     ' Copy String Value
                Case REG_DWORD                                          ' Double Word Registry Key Data Type
                    For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                    Next
                    KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
            End Select

            result = True                                      ' Return Success
            rc = RegCloseKey(hKey)                                  ' Close Registry Key

        Catch ex As Exception
            ' Cleanup After An Error Has Occured...
            KeyVal = ""                                             ' Set Return Val To Empty String
            result = False                                     ' Return Failure
            rc = RegCloseKey(hKey)                                  ' Close Registry Key

            ExHandle(ex)
            Return result

        End Try

        Return result

    End Function

    Private Sub WriteObjectClassInformation()

        Try

            Dim pDataset As IDataset
            Dim pFeatureWorkspace As IFeatureWorkspace
            Dim pDatasetSplit As Object
            Dim pCounter As Long
            Dim pTable As Table
            Dim pFeatureClass As IFeatureClass
            '
            pFeatureWorkspace = mWorkspace

            For pCounter = 0 To Me.lstDataset.Items.Count - 1 Step 1

                pDatasetSplit = Split(Me.lstDataset.Items.Item(pCounter), "|")

                Select Case CLng(pDatasetSplit(3))

                    Case esriDatasetType.esriDTTable
                        pTable = pFeatureWorkspace.OpenTable(CStr(pDatasetSplit(2)))
                        pDataset = pTable
                        Call WriteObjectClassInformation2(pDataset)

                    Case esriDatasetType.esriDTFeatureClass
                        pFeatureClass = pFeatureWorkspace.OpenFeatureClass(CStr(pDatasetSplit(2)))
                        pDataset = pFeatureClass
                        Call WriteObjectClassInformation2(pDataset)
                    Case Else
                        '
                End Select
            Next pCounter


        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub WriteObjectClassInformation2(ByRef pDataset As IDataset)

        Dim pObjectClass As IObjectClass
        Dim pSubtypes As ISubtypes
#Disable Warning BC42024 ' Unused local variable: 'pSubtypeCounter'.
        Dim pSubtypeCounter As Long
#Enable Warning BC42024 ' Unused local variable: 'pSubtypeCounter'.
#Disable Warning BC42024 ' Unused local variable: 'pSubtypeName'.
        Dim pSubtypeName As String
#Enable Warning BC42024 ' Unused local variable: 'pSubtypeName'.
#Disable Warning BC42024 ' Unused local variable: 'pSubtypeString'.
        Dim pSubtypeString As String
#Enable Warning BC42024 ' Unused local variable: 'pSubtypeString'.
        Dim pFeatureClass As IFeatureClass

        Dim lGeomIndex As Long
        Dim sShpName As String
        Dim pShapeField As IField
        Dim pGeometryDef As IGeometryDef
        Dim sHasMs As String
        Dim sHasZs As String

        Dim pDomain As IDomain

        Dim sFieldArray(100) As Boolean
        Dim numFieldArray As Long
        Dim i As Long
#Disable Warning BC42024 ' Unused local variable: 'j'.
        Dim j As Long
#Enable Warning BC42024 ' Unused local variable: 'j'.

#Disable Warning BC42024 ' Unused local variable: 'sDomain'.
        Dim sDomain As String
#Enable Warning BC42024 ' Unused local variable: 'sDomain'.
#Disable Warning BC42024 ' Unused local variable: 'sDefaultValue'.
        Dim sDefaultValue As String
#Enable Warning BC42024 ' Unused local variable: 'sDefaultValue'.
#Disable Warning BC42024 ' Unused local variable: 'indexDomain'.
        Dim indexDomain As Long
#Enable Warning BC42024 ' Unused local variable: 'indexDomain'.
#Disable Warning BC42024 ' Unused local variable: 'indexDefaultValue'.
        Dim indexDefaultValue As Long
#Enable Warning BC42024 ' Unused local variable: 'indexDefaultValue'.

        Dim pEnumSubTypes As IEnumSubtype
        Dim lSubT As Long
        Dim sSubT As String
        Dim sFieldName As String
        Dim numSubtypes As Long

        Dim sSubtypeField As String
        Dim sDefaultSubtype As String
        Dim sSubtypeName As String
        Dim sSubtypeCode As String

        Dim sFieldNames As String
        Dim sDefaultValues As String
        Dim sDomains As String
        Dim sClassName As String

        Dim lSubtypeFieldIndex As Long
        Dim lSubtypeFromBottom As Long

        Dim pMetadata As IMetadata
        Dim pPropSet As IPropertySet
        Dim vAbstractProp As Object
        Dim sAbstract As String

        pPropSet = Nothing

        Try

            pObjectClass = pDataset
            If pDataset.Type = esriDatasetType.esriDTFeatureClass Then
                pFeatureClass = pDataset
                pMetadata = pFeatureClass 'QI
                pPropSet = pMetadata.Metadata

                ' Metadata Abstract
                If Me.chkAbstract.Checked Then
                    'Set pMetadata = pFeatureClass 'QI
                    'Set pPropSet = pMetadata.metadata
                    vAbstractProp = pPropSet.GetProperty("idinfo/descript/abstract")
                    If vAbstractProp Is Nothing Then
                        sAbstract = "No metadata abstract"
                    Else
                        sAbstract = CStr(vAbstractProp(0))
                    End If

                Else
                    sAbstract = ""
                End If


                'Go to IGeometryDef to determine whether the feature class has z's and m's
                sShpName = pFeatureClass.ShapeFieldName
                lGeomIndex = pFeatureClass.Fields.FindField(sShpName)
                pShapeField = pFeatureClass.Fields.Field(lGeomIndex)
                pGeometryDef = pShapeField.GeometryDef
                If pGeometryDef.HasM = True Then
                    sHasMs = "Yes"
                Else
                    sHasMs = "No"
                End If
                If pGeometryDef.HasZ = True Then
                    sHasZs = "Yes"
                Else
                    sHasZs = "No"
                End If

                VisioDiagram.StartFeatureClass(pDataset.Name, pFeatureClass.FeatureType,
                  pFeatureClass.ShapeType, pFeatureClass.Fields.FieldCount, sHasMs, sHasZs, sAbstract)

            Else
                '-------------------------
                ' The Dataset is a Table.
                '-------------------------
                pObjectClass = pDataset
                pMetadata = pObjectClass 'QI
                pPropSet = pMetadata.Metadata

                ' Metadata Abstract
                If Me.chkAbstract.Checked Then
                    'Set pMetadata = pObjectClass 'QI
                    'Set pPropSet = pMetadata.metadata
                    vAbstractProp = pPropSet.GetProperty("idinfo/descript/abstract")
                    sAbstract = CStr(vAbstractProp(0))
                Else
                    sAbstract = ""
                End If

                VisioDiagram.startTable(pDataset.Name, pObjectClass.Fields.FieldCount, sAbstract)
            End If
            '
            pSubtypes = pObjectClass

            'Here, we write out the fields regardless of whether there are subtypes or not. If subtypes, then the first subtype gets
            'written out. If multiple subtypes, then a difference table is generated. MNZ

            ' Option to Omit the fields for annotation featureclasses.
            If Me.chkOmitAnno.Checked Then
                If pDataset.Type = esriDatasetType.esriDTFeatureClass Then
#Disable Warning BC42104 ' Variable 'pFeatureClass' is used before it has been assigned a value. A null reference exception could result at runtime.
                    If pFeatureClass.FeatureType = esriFeatureType.esriFTAnnotation Then
#Enable Warning BC42104 ' Variable 'pFeatureClass' is used before it has been assigned a value. A null reference exception could result at runtime.

                        Dim pFieldsEmpty As IFields
                        pFieldsEmpty = New Fields
                        Dim pFieldsEdit As IFieldsEdit
                        pFieldsEdit = pFieldsEmpty
                        pFieldsEdit.AddField(pObjectClass.Fields.Field(pObjectClass.Fields.FindField("OBJECTID")))

                        WriteObjectClassInformation3(pFieldsEmpty, pDataset.Name, metadata:=pPropSet)
                    Else
                        WriteObjectClassInformation3(pObjectClass.Fields, pDataset.Name, metadata:=pPropSet)
                    End If
                Else
                    WriteObjectClassInformation3(pObjectClass.Fields, pDataset.Name, metadata:=pPropSet)
                End If
            Else
                WriteObjectClassInformation3(pObjectClass.Fields, pDataset.Name, metadata:=pPropSet)
            End If


            'Here comes the subtypes difference table...

            If pSubtypes.HasSubtype Then

                'First, go through all subtypes, inspect default values and domains, and build list of fields which have non-null values
                For i = 0 To 99
                    sFieldArray(i) = False
                Next i

                '  get  the enumeration of all of the subtypes for this feature class

                pEnumSubTypes = pSubtypes.Subtypes

                ' loop through all of the subtypes and bring up a message
                ' box with each SubType 's code and name
                lSubT = 0
                numSubtypes = 0
                sSubT = pEnumSubTypes.Next(lSubT)
                Dim initSubT As Long
                Dim lastSubT As Long
                initSubT = lSubT
                lastSubT = 0

                Do While sSubT <> ""

                    For i = 0 To pObjectClass.Fields.FieldCount - 1
                        sFieldName = pObjectClass.Fields.Field(i).Name

                        If UCase(sFieldName) <> "OBJECTID" And UCase(sFieldName) <> "SHAPE" Then
                            If pSubtypes.DefaultValue(lSubT, sFieldName) Is Nothing Then
                                sFieldArray(i) = True
                            End If
                        End If

                        If UCase(sFieldName) <> "OBJECTID" And UCase(sFieldName) <> "SHAPE" Then
                            pDomain = pSubtypes.Domain(lSubT, sFieldName)

                            If pDomain Is Nothing Then
                            Else
                                If pDomain.Name <> "" Then
                                    sFieldArray(i) = True
                                End If
                            End If
                        End If

                    Next i
                    'MsgBox lSubT & ": " & pSubtypes.SubtypeName(lSubT)
                    sSubT = pEnumSubTypes.Next(lSubT)
                    numSubtypes = numSubtypes + 1
                Loop

                'Now write out subtype summary underneath the table/feature class...
                sSubtypeField = pSubtypes.SubtypeFieldName
                sDefaultSubtype = pSubtypes.DefaultSubtypeCode

                'Count how many fields are not null
                numFieldArray = 0
                For i = 0 To pObjectClass.Fields.FieldCount - 1
                    If sFieldArray(i) = True Then numFieldArray = numFieldArray + 1
                    If UCase(pObjectClass.Fields.Field(i).Name) = UCase(sSubtypeField) Then
                        lSubtypeFieldIndex = i
                    End If
                Next i

                'Write subtype header

                lSubtypeFromBottom = pObjectClass.Fields.FieldCount - lSubtypeFieldIndex
                sClassName = pDataset.Name
                VisioDiagram.startSubtype(sClassName, sSubtypeField, lSubtypeFromBottom, sDefaultSubtype, numSubtypes, numFieldArray)

                'Loop through subtypes and write out subtype default value and domain name

                pEnumSubTypes = pSubtypes.Subtypes

                lSubT = 0
                pEnumSubTypes.Reset()
                'sSubT = pSubtypes.SubtypeName(lSubT)

                lSubT = 0
                sSubT = pEnumSubTypes.Next(lSubT)

                Do While sSubT <> ""

                    sFieldNames = ""
                    sDefaultValues = ""
                    sDomains = ""

                    'Get subtype code and description
                    sSubtypeName = pSubtypes.SubtypeName(lSubT)
                    sSubtypeCode = CStr(lSubT)

                    For i = 0 To pObjectClass.Fields.FieldCount - 1
                        If sFieldArray(i) = True Then

                            sFieldName = pObjectClass.Fields.Field(i).Name
                            sFieldNames = sFieldNames & sFieldName & ":"
                            sDefaultValues = sDefaultValues & pSubtypes.DefaultValue(lSubT, sFieldName) & ":"
                            If pSubtypes.Domain(lSubT, sFieldName) Is Nothing Then
                                sDomains = sDomains & ":"
                            Else
                                sDomains = sDomains & pSubtypes.Domain(lSubT, sFieldName).Name & ":"
                            End If

                        End If
                    Next i

                    VisioDiagram.diagramSubtype(sSubtypeCode, sSubtypeName, numFieldArray, sFieldNames, sDefaultValues, sDomains)
                    sSubT = pEnumSubTypes.Next(lSubT)

                Loop

                VisioDiagram.finishSubtype()

            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Private Sub WriteDomainInformation()
        '

        Dim pIndexDomain As Long
        Dim pIndexDataset As Long
        Dim pDomain As IDomain
#Disable Warning BC42024 ' Unused local variable: 'pRangeDomain'.
        Dim pRangeDomain As IRangeDomain
#Enable Warning BC42024 ' Unused local variable: 'pRangeDomain'.
        Dim pCodedValueDomain As ICodedValueDomain
        Dim pWorkspaceDomains As IWorkspaceDomains
#Disable Warning BC42024 ' Unused local variable: 'pFeatureClassContainer'.
        Dim pFeatureClassContainer As IFeatureClassContainer
#Enable Warning BC42024 ' Unused local variable: 'pFeatureClassContainer'.
#Disable Warning BC42024 ' Unused local variable: 'pEnumFeatureClass'.
        Dim pEnumFeatureClass As IEnumFeatureClass
#Enable Warning BC42024 ' Unused local variable: 'pEnumFeatureClass'.
#Disable Warning BC42024 ' Unused local variable: 'pFeatureClass'.
        Dim pFeatureClass As IFeatureClass
#Enable Warning BC42024 ' Unused local variable: 'pFeatureClass'.
        Dim pFeatureWorkspace As IFeatureWorkspace
        Dim pSubtypes As ISubtypes
        Dim pEnumSubtype As IEnumSubtype
        Dim pObjectClass As IObjectClass
        Dim pDataset As IDataset
        'Dim pDomainName As String
        Dim pSubtypeName As String
#Disable Warning BC42024 ' Unused local variable: 'pIndexDM'.
        Dim pIndexDM As Long
#Enable Warning BC42024 ' Unused local variable: 'pIndexDM'.
        Dim pIndexFD As Long
        Dim pIndexST As Long
        Dim pSplitDS As Object
        Dim pSplitDM As Object
        Dim pAssignedToObjectClass As Boolean

        Dim numDomainValues As Long

        Try
            '--------------------
            ' Write Table Header
            '--------------------
            '
            If Me.lstDomain.Items.Count <> 0 Then
                pWorkspaceDomains = mWorkspace
                pFeatureWorkspace = mWorkspace
                '----------------------------------------------------------------------------------------------
                ' To improve performance lets load all the ObjectClasses/Subtypes/Fields and assigned Domains.
                ' This will make Domain-To-Field searching faster!
                ' Temp ListBox (Me.lstTemp) format:
                ' DomainName | DatasetName | DatasetType | SubtypeName | FieldName
                '----------------------------------------------------------------------------------------------
                Me.lstTemp.Items.Clear()


                For pIndexDataset = 0 To Me.lstDataset.Items.Count - 1 Step 1
                    pSplitDS = Split(Me.lstDataset.Items.Item(pIndexDataset), "|")
                    If CLng(pSplitDS(3)) = esriDatasetType.esriDTFeatureClass Or
                       CLng(pSplitDS(3)) = esriDatasetType.esriDTTable Then
                        pDataset = pFeatureWorkspace.OpenTable(CStr(pSplitDS(2)))
                        pObjectClass = pDataset
                        pSubtypes = pObjectClass
                        If pSubtypes.HasSubtype Then
                            pEnumSubtype = pSubtypes.Subtypes
                            pSubtypeName = pEnumSubtype.Next(pIndexST)
                            Do Until pSubtypeName = ""
                                For pIndexFD = 0 To pObjectClass.Fields.FieldCount - 1 Step 1
                                    pDomain = pSubtypes.Domain(pIndexST, pObjectClass.Fields.Field(pIndexFD).Name)
                                    If Not (pDomain Is Nothing) Then
                                        Me.lstTemp.Items.Add(pDomain.Name & "|" &
                                                           CStr(pSplitDS(2)) & "|" &
                                                           CLng(pSplitDS(3)) & "|" &
                                                           pSubtypeName & "|" &
                                                           pObjectClass.Fields.Field(pIndexFD).Name)

                                    End If
                                Next
                                pSubtypeName = pEnumSubtype.Next(pIndexST)
                            Loop
                        Else
                            For pIndexFD = 0 To pObjectClass.Fields.FieldCount - 1 Step 1
                                pDomain = pObjectClass.Fields.Field(pIndexFD).Domain
                                If Not (pDomain Is Nothing) Then
                                    Me.lstTemp.Items.Add(pDomain.Name & "|" &
                                                       CStr(pSplitDS(2)) & "|" &
                                                       CLng(pSplitDS(3)) & "|" &
                                                       mNoData & "|" &
                                                       pObjectClass.Fields.Field(pIndexFD).Name)

                                End If
                            Next
                        End If
                    End If
                Next pIndexDataset
                '
                For pIndexDomain = 0 To Me.lstDomain.Items.Count - 1 Step 1
                    pSplitDM = Split(Me.lstDomain.Items.Item(pIndexDomain), "|")
                    pDomain = pWorkspaceDomains.DomainByName(CStr(pSplitDM(0)))
                    pAssignedToObjectClass = False

                    '-------------
                    ' Domain Name
                    '-------------

                    'First, find the number of domain values so that can be passed to ModVisioDiagram to size the box properly
                    Select Case CLng(pSplitDM(1))
                        Case esriDomainType.esriDTCodedValue
                            pCodedValueDomain = pDomain
                            numDomainValues = pCodedValueDomain.CodeCount
                        Case esriDomainType.esriDTRange
                            numDomainValues = 1
                    End Select

                    VisioDiagram.startDomain(pDomain.Name,
                                            pDomain.Description,
                                            pDomain.FieldType,
                                            CLng(pSplitDM(1)),
                                            pDomain.MergePolicy,
                                            pDomain.SplitPolicy,
                                            numDomainValues, pDomain)

                Next pIndexDomain
            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub

    Private Sub WriteRelationshipClassInformation()

        Try

            Dim pCounter As Long
#Disable Warning BC42024 ' Unused local variable: 'pCounter2'.
            Dim pCounter2 As Long
#Enable Warning BC42024 ' Unused local variable: 'pCounter2'.
            Dim pDatasetSplit As Object
            Dim pFeatureWorkspace As IFeatureWorkspace
#Disable Warning BC42024 ' Unused local variable: 'pSubtypeName'.
            Dim pSubtypeName As String
#Enable Warning BC42024 ' Unused local variable: 'pSubtypeName'.
#Disable Warning BC42024 ' Unused local variable: 'pEnumSubtype'.
            Dim pEnumSubtype As IEnumSubtype
#Enable Warning BC42024 ' Unused local variable: 'pEnumSubtype'.
#Disable Warning BC42024 ' Unused local variable: 'pSubtypesOrigin'.
            Dim pSubtypesOrigin As ISubtypes
#Enable Warning BC42024 ' Unused local variable: 'pSubtypesOrigin'.
#Disable Warning BC42024 ' Unused local variable: 'pSubtypesDestination'.
            Dim pSubtypesDestination As ISubtypes
#Enable Warning BC42024 ' Unused local variable: 'pSubtypesDestination'.
#Disable Warning BC42024 ' Unused local variable: 'pRelationshipRule'.
            Dim pRelationshipRule As IRelationshipRule
#Enable Warning BC42024 ' Unused local variable: 'pRelationshipRule'.
#Disable Warning BC42024 ' Unused local variable: 'pEnumRule'.
            Dim pEnumRule As IEnumRule
#Enable Warning BC42024 ' Unused local variable: 'pEnumRule'.
#Disable Warning BC42024 ' Unused local variable: 'lngRule'.
            Dim lngRule As Long
#Enable Warning BC42024 ' Unused local variable: 'lngRule'.
            Dim pDatasetOrigin As IDataset
            Dim pDatasetDestination As IDataset
            Dim pRelationshipClass As IRelationshipClass
            Dim pRelationshipTable As ITable
            Dim pDataset As IDataset

            Dim cRelationshipType As String
            Dim cCardinality As String
            Dim cNotification As String
            Dim cBackwardLabel As String
            Dim cForwardLabel As String

            Dim cAttributed As String

            Dim cOriginName As String
            Dim iOriginType As Integer
            Dim cOriginPrimaryKey As String
            Dim cOriginForeignKey As String

            Dim cDestinationName As String
            Dim iDestinationType As Integer
            Dim cDestinationPrimaryKey As String
            Dim cDestinationForeignKey As String

#Disable Warning BC42024 ' Unused local variable: 'relationshipFields'.
            Dim relationshipFields As Fields
#Enable Warning BC42024 ' Unused local variable: 'relationshipFields'.

            Dim lIsManyToMany As Boolean
            Dim cRelProp1 As String
            Dim cRelProp2 As String
            Dim cRelProp3 As String
            Dim cRelProp4 As String

            '----------------------------------
            ' Check if any Relationships Exist
            '----------------------------------
            Me.lstTemp.Items.Clear()
            For pCounter = 0 To Me.lstDataset.Items.Count - 1 Step 1
                pDatasetSplit = Split(Me.lstDataset.Items.Item(pCounter), "|")
                If CLng(pDatasetSplit(3)) = esriDatasetType.esriDTRelationshipClass Then
                    Me.lstTemp.Items.Add(pDatasetSplit(2))
                End If
            Next pCounter
            '
            If Me.lstTemp.Items.Count <> 0 Then
                pFeatureWorkspace = mWorkspace
                For pCounter = 0 To Me.lstTemp.Items.Count - 1 Step 1
                    pRelationshipClass = pFeatureWorkspace.OpenRelationshipClass(Me.lstTemp.Items.Item(pCounter))
                    pDataset = pRelationshipClass

                    If pRelationshipClass.IsComposite Then
                        cRelationshipType = "Composite"
                    Else
                        cRelationshipType = "Simple"
                    End If
                    '--------------------------
                    ' Relationship Cardinality
                    '--------------------------
                    Select Case pRelationshipClass.Cardinality
                        Case esriRelCardinality.esriRelCardinalityOneToOne
                            cCardinality = "One to one"
                        Case esriRelCardinality.esriRelCardinalityOneToMany
                            cCardinality = "One to many"
                        Case esriRelCardinality.esriRelCardinalityManyToMany
                            cCardinality = "Many to many"
                    End Select

                    '---------------------------
                    ' Relationship Notification
                    '---------------------------
                    Select Case pRelationshipClass.Notification
                        Case esriRelNotification.esriRelNotificationNone
                            cNotification = "None"
                        Case esriRelNotification.esriRelNotificationForward
                            cNotification = "Forward"
                        Case esriRelNotification.esriRelNotificationBackward
                            cNotification = "Backward"
                        Case esriRelNotification.esriRelNotificationBoth
                            cNotification = "Both"
                    End Select

                    '--------------------------
                    ' Relationship Attributed?
                    '--------------------------
                    If pRelationshipClass.IsAttributed Then
                        cAttributed = "Attributed"
                        cOriginForeignKey = pRelationshipClass.OriginForeignKey
                        cDestinationForeignKey = pRelationshipClass.DestinationForeignKey
                        pRelationshipTable = pDataset

                    Else
                        cAttributed = "Not attributed"
                    End If
                    '------------------------------------------
                    ' Origin and Destination ObjectClass Names
                    '------------------------------------------
                    pDatasetOrigin = pRelationshipClass.OriginClass

                    cOriginName = pDatasetOrigin.Name
                    iOriginType = pDatasetOrigin.Type

                    pDatasetDestination = pRelationshipClass.DestinationClass

                    cDestinationName = pDatasetDestination.Name
                    iDestinationType = pDatasetDestination.Type

                    '-----------------------------
                    ' Origin and Destination Keys
                    '-----------------------------
                    cOriginPrimaryKey = pRelationshipClass.OriginPrimaryKey
                    cOriginForeignKey = pRelationshipClass.OriginForeignKey

                    cDestinationPrimaryKey = pRelationshipClass.DestinationPrimaryKey
                    cDestinationForeignKey = pRelationshipClass.DestinationForeignKey

                    '------------------------------------------
                    ' Forward and Backward Relationship Labels
                    '------------------------------------------
                    cForwardLabel = pRelationshipClass.ForwardPathLabel
                    cBackwardLabel = pRelationshipClass.BackwardPathLabel

                    '------------------------------------------
                    ' Diagram relationship class so far...
                    '------------------------------------------
#Disable Warning BC42104 ' Variable 'cCardinality' is used before it has been assigned a value. A null reference exception could result at runtime.
#Disable Warning BC42104 ' Variable 'cNotification' is used before it has been assigned a value. A null reference exception could result at runtime.
                    cRelProp1 = cRelationshipType & vbCrLf & cCardinality & vbCrLf & cNotification
#Enable Warning BC42104 ' Variable 'cNotification' is used before it has been assigned a value. A null reference exception could result at runtime.
#Enable Warning BC42104 ' Variable 'cCardinality' is used before it has been assigned a value. A null reference exception could result at runtime.
                    cRelProp2 = cForwardLabel & vbCrLf & cBackwardLabel
                    cRelProp3 = cOriginName & vbCrLf & cOriginPrimaryKey & vbCrLf & cOriginForeignKey
                    cRelProp4 = cDestinationName

                    If cCardinality = "Many to many" Then
                        lIsManyToMany = True
                        cRelProp4 = cDestinationName & vbCrLf & cDestinationPrimaryKey & vbCrLf & cDestinationForeignKey
                    Else
                        lIsManyToMany = False
                        cRelProp4 = cDestinationName
                    End If

                    If pRelationshipClass.IsAttributed Then
#Disable Warning BC42104 ' Variable 'pRelationshipTable' is used before it has been assigned a value. A null reference exception could result at runtime.
                        VisioDiagram.DiagramRelationship(pDataset.Name, lIsManyToMany, cRelProp1, cRelProp2, cRelProp3, cRelProp4,
                          iOriginType, iDestinationType, pRelationshipClass.IsAttributed, pRelationshipTable.Fields)
#Enable Warning BC42104 ' Variable 'pRelationshipTable' is used before it has been assigned a value. A null reference exception could result at runtime.
                    Else
                        VisioDiagram.DiagramRelationship(pDataset.Name, lIsManyToMany, cRelProp1, cRelProp2, cRelProp3, cRelProp4,
                         iOriginType, iDestinationType, pRelationshipClass.IsAttributed)
                    End If

                Next pCounter

            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Private Function ValidateSubtypeCode(ByRef pSubtypes As ISubtypes, ByRef pSubtypeCodeIn As Long) As Long
        '
        Dim result As Long = 0
        Try

            ValidateSubtypeCode = False

            Dim pEnumSubtype As IEnumSubtype
            Dim pSubtypeName As String
            Dim pSubtypeCodeOut As Long

            pEnumSubtype = pSubtypes.Subtypes
            pSubtypeName = pEnumSubtype.Next(pSubtypeCodeOut)
            Do Until pSubtypeName = ""
                If pSubtypeCodeIn = pSubtypeCodeOut Then
                    result = True
                    Exit Do
                End If
                pSubtypeName = pEnumSubtype.Next(pSubtypeCodeOut)
            Loop

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result
    End Function

    Private Function ListBoxRowCount(ByRef pValue As String) As Long
        '
        Dim result As Long = 0
        Try

            Dim pIndex As Long
            Dim pRowCount As Long
#Disable Warning BC42024 ' Unused local variable: 'pValueArray'.
            Dim pValueArray As Object
#Enable Warning BC42024 ' Unused local variable: 'pValueArray'.
#Disable Warning BC42024 ' Unused local variable: 'pValueIndex'.
            Dim pValueIndex As Long
#Enable Warning BC42024 ' Unused local variable: 'pValueIndex'.
            '
            pRowCount = 0
            For pIndex = 0 To Me.lstTemp.Items.Count - 1 Step 1
                If Strings.Left(Me.lstTemp.Items.Item(pIndex), Len(pValue) + 1) = pValue & "|" Then
                    pRowCount = pRowCount + 1
                End If
            Next pIndex
            result = pRowCount

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result

    End Function

    Private Sub ClearModularVariables()
        '
        Try

            Me.lstDataset.Items.Clear()
            Me.lstDomain.Items.Clear()
            Me.lstTemp.Items.Clear()

            ' ReDim mLanguageReference(0, 0) As String

            mWorkspace = Nothing
            mGxApplication = Nothing

        Catch ex As Exception
            ExHandle(ex)
        End Try

    End Sub

    Private Function BooleanToYesNo(ByRef pBoolean As Boolean) As String
        ' Converts a Boolean value to English String.

        Dim result As String = ""
        Try
            If pBoolean Then
                result = "Yes"
            Else
                result = "No"
            End If

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Return result
    End Function

#End Region

#Region "Public Functions"

    Private Sub WriteObjectClassInformation3(ByRef pFields As IFields,
                                        ByRef pDatasetName As String,
                                        Optional ByRef pSubtypes As ISubtypes = Nothing,
                                        Optional ByRef lngSubtypeCode As Long = 0,
                                        Optional ByVal metadata As IPropertySet = Nothing)
        Try

            Dim pFieldCounter As Long
            Dim pDomain As IDomain

            Call VisioDiagram.DiagramFeatureClass(pFields, pDatasetName, metadata)

            For pFieldCounter = 0 To pFields.FieldCount - 1 Step 1
                If pSubtypes Is Nothing Then
                    pDomain = pFields.Field(pFieldCounter).Domain
                Else
                    pDomain = pSubtypes.Domain(lngSubtypeCode, pFields.Field(pFieldCounter).Name)
                End If
            Next

        Catch ex As Exception
            ExHandle(ex)
        End Try

        Exit Sub
    End Sub

#End Region

End Class