Imports System
Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Reflection

Public Class LocalizationManager
    Private Shared _resourceManager As ResourceManager
    Private Shared _currentCulture As CultureInfo
    
    ''' <summary>
    ''' Initialize the localization manager with the current Outlook UI culture
    ''' </summary>
    ''' <param name="outlookUICulture">The culture from Outlook UI, or Nothing to use system default</param>
    Public Shared Sub Initialize(Optional outlookUICulture As CultureInfo = Nothing)
        ' Initialize the resource manager
        _resourceManager = New ResourceManager("OutlookDLP.Strings", Assembly.GetExecutingAssembly())
        
        ' Set the current culture based on Outlook UI culture or system settings
        If outlookUICulture IsNot Nothing Then
            _currentCulture = outlookUICulture
        Else
            ' Try to get the current UI culture from Outlook
            Try
                Dim outlookApp = Globals.ThisAddIn.Application
                Dim languageSettings = outlookApp.LanguageSettings
                
                ' Get the language ID for the UI
                Dim lcid As Integer = languageSettings.LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI)
                _currentCulture = New CultureInfo(lcid)
            Catch ex As Exception
                ' Fallback to current thread culture
                _currentCulture = Thread.CurrentThread.CurrentUICulture
            End Try
        End If
        
        ' Set the current UI culture
        Thread.CurrentThread.CurrentUICulture = _currentCulture
    End Sub
    
    ''' <summary>
    ''' Get a localized string
    ''' </summary>
    ''' <param name="key">The resource key</param>
    ''' <returns>The localized string</returns>
    Public Shared Function GetString(key As String) As String
        Try
            Return _resourceManager.GetString(key, _currentCulture)
        Catch ex As Exception
            ' Fallback to the key itself if the resource is not found
            Return key
        End Try
    End Function
    
    ''' <summary>
    ''' Get a localized string with format parameters
    ''' </summary>
    ''' <param name="key">The resource key</param>
    ''' <param name="args">Format arguments</param>
    ''' <returns>The formatted localized string</returns>
    Public Shared Function GetString(key As String, ParamArray args() As Object) As String
        Try
            Dim format As String = _resourceManager.GetString(key, _currentCulture)
            Return String.Format(format, args)
        Catch ex As Exception
            ' Fallback to the key itself if the resource is not found
            Return key
        End Try
    End Function
    
    ''' <summary>
    ''' Get the current culture
    ''' </summary>
    ''' <returns>The current culture</returns>
    Public Shared Function GetCurrentCulture() As CultureInfo
        Return _currentCulture
    End Function
    
    ''' <summary>
    ''' Set the current culture
    ''' </summary>
    ''' <param name="culture">The culture to set</param>
    Public Shared Sub SetCulture(culture As CultureInfo)
        _currentCulture = culture
        Thread.CurrentThread.CurrentUICulture = culture
    End Sub
End Class