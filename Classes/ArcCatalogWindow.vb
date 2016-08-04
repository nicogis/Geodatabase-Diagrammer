Imports System
Imports ESRI.ArcGIS.Framework

Public Class ArcCatalogWindow
    Implements System.Windows.Forms.IWin32Window
#Region "Private Variables"
    Private m_app As ESRI.ArcGIS.Framework.IApplication
#End Region

#Region "Constructor"

    Public Sub New(ByVal application As ESRI.ArcGIS.Framework.IApplication)
        m_app = application
    End Sub
#End Region

#Region "Public Properties"
    Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
        Get
            Return New IntPtr(m_app.hWnd)
        End Get
    End Property
#End Region
End Class
