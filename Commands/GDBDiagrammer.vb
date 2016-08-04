Imports System.Windows.Forms.Form


Public Class GDBDiagrammer
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
#Region "Public Variables"


#End Region

#Region "Member Variables"

#End Region

#Region "Event Handlers"
    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        ' Open the main form.
        Try
            Dim pform As FormGDBDiagrammer = New FormGDBDiagrammer
            pform.ShowDialog(New ArcCatalogWindow(My.ArcCatalog.Application))

            ' Returns back to here when the form is closed.

        Catch ex As Exception
            ExHandle(ex)
        End Try

        'My.ArcCatalog.Application.CurrentTool = Nothing
    End Sub

    Protected Overrides Sub OnUpdate()
        ' Enable if a geodatabase is the selected object.

        Try

            Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject = My.ThisApplication.SelectedObject
            If TypeOf pGxObject Is ESRI.ArcGIS.Catalog.IGxDatabase Then
                Enabled = True
            Else
                Enabled = False
            End If
            'Enabled = My.ArcCatalog.Application IsNot Nothing

        Catch ex As Exception
            ExHandle(ex)
        End Try
    End Sub
#End Region

End Class
