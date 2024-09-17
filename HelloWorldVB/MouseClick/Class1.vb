Imports System.Reflection
Imports System.Windows.Forms
Imports Autodesk.Connectivity.Explorer.Extensibility
Imports Autodesk.Connectivity.WebServicesTools
Imports Autodesk.Connectivity.WebServices


' These 5 assembly attributes must be specified or your extension will not load. 
<Assembly: AssemblyCompany("Autodesk")>
<Assembly: AssemblyProduct("MouseClickVaultExtension")>
<Assembly: AssemblyDescription("MouseClick App")>

' The extension ID needs to be unique for each extension.  
' Make sure to generate your own ID when writing your own extension. 
<Assembly: Autodesk.Connectivity.Extensibility.Framework.ExtensionId("7ADC0766-F085-46d7-A2EB-C68F79CBF4E8")>

' This number gets incremented for each Vault release.
<Assembly: Autodesk.Connectivity.Extensibility.Framework.ApiVersion("17.0")>

Namespace VaultKeyboardEventExtension

    ''' <summary>
    ''' This class implements the IExtension interface, which means it tells Vault Explorer what 
    ''' commands and custom tabs are provided by this extension.
    ''' </summary>
    Public Class MouseClickVaultExtension
        Implements IExplorerExtension

        Private connection As Connection
        Private interactionEvents As InteractionEvents
        Private keyPressHandler As KeyboardEventsSink_OnKeyPressEventHandler

        Public Sub OnLogOff(ByVal application As IApplication)
            UnregisterEvents()
        End Sub

        Public Sub OnLogOn(ByVal application As IApplication)
            connection = CType(application.Connection, Connection)
            RegisterEvents()
        End Sub

        Public Sub OnStartup(ByVal application As IApplication)
        End Sub

        Public Sub OnShutdown(ByVal application As IApplication)
            UnregisterEvents()
        End Sub

        Private Sub RegisterEvents()
            If connection IsNot Nothing Then
                interactionEvents = connection.InteractionManager.CreateInteractionEvents()
                keyPressHandler = New KeyboardEventsSink_OnKeyPressEventHandler(AddressOf OnKeyPress)
                interactionEvents.KeyboardEvents.OnKeyPress += keyPressHandler
                interactionEvents.Start()
            End If
        End Sub

        Private Sub UnregisterEvents()
            If interactionEvents IsNot Nothing Then
                interactionEvents.KeyboardEvents.OnKeyPress -= keyPressHandler
                interactionEvents.[Stop]()
                interactionEvents = Nothing
            End If
        End Sub

        Private Sub OnKeyPress(ByVal keyAscii As Integer)
            If keyAscii = 38 OrElse keyAscii = 40 Then
                MessageBox.Show("Eureka", "Key Press Detected", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Public ReadOnly Property ExtensionInfo As ExtensionInfo
            Get
                Return New ExtensionInfo With {
                    .Name = "Vault Keyboard Event Extension",
                    .Description = "Displays a message when arrow keys are pressed",
                    .Version = "1.0",
                    .Author = "YourName"
                }
            End Get
        End Property
    End Class
End Namespace
