Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms

Imports Autodesk.Connectivity.Explorer.Extensibility
Imports Autodesk.Connectivity.WebServices
Imports Autodesk.Connectivity.WebServicesTools
Imports VDF = Autodesk.DataManagement.Client.Framework


' These 5 assembly attributes must be specified or your extension will not load. 
<Assembly: AssemblyCompany("Autodesk")>
<Assembly: AssemblyProduct("HelloWorldCommandExtension")>
<Assembly: AssemblyDescription("Sample App")>

' The extension ID needs to be unique for each extension.  
' Make sure to generate your own ID when writing your own extension. 
<Assembly: Autodesk.Connectivity.Extensibility.Framework.ExtensionId("7ADC0766-F085-46d7-A2EB-C68F79CBF4E7")>

' This number gets incremented for each Vault release.
<Assembly: Autodesk.Connectivity.Extensibility.Framework.ApiVersion("18.0")>

Namespace HelloWorld

    ''' <summary>
    ''' This class implements the IExtension interface, which means it tells Vault Explorer what 
    ''' commands and custom tabs are provided by this extension.
    ''' </summary>
    Public Class HelloWorldCommandExtension
        Implements IExplorerExtension

#Region "IExtension Members"

        ''' <summary>
        ''' This function tells Vault Explorer what custom commands this extension provides.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <returns>A collection of CommandSites, which are collections of custom commands.</returns>
        Public Function CommandSites() As IEnumerable(Of CommandSite) Implements IExplorerExtension.CommandSites
            ' Create the Hello World command object.
            ' this command is active when a File is selected

            ' this command is not active if there are multiple entities selected
            Dim helloWorldCmdItem As New CommandItem("HelloWorldCommand", "Visor") With {
                .NavigationTypes = New SelectionTypeId() {SelectionTypeId.File, SelectionTypeId.FileVersion, SelectionTypeId.Item, SelectionTypeId.Bom},
                .MultiSelectEnabled = False
            }

            ' The HelloWorldCommandHandler function is called when the custom command is executed.
            AddHandler helloWorldCmdItem.Execute, AddressOf HelloWorldCommandHandler

            ' Create a command site to hook the command to the Advanced toolbar
            Dim toolbarCmdSite As New CommandSite("HelloWorldCommand.Toolbar", "Visor") With {
             .Location = CommandSiteLocation.AdvancedToolbar,
             .DeployAsPulldownMenu = False
            }
            toolbarCmdSite.AddCommand(helloWorldCmdItem)

            ' Create another command site to hook the command to the right-click menu for Files.
            Dim fileContextCmdSite As New CommandSite("HelloWorldCommand.FileContextMenu", "Visor") With {
             .Location = CommandSiteLocation.FileContextMenu,
             .DeployAsPulldownMenu = False
            }
            fileContextCmdSite.AddCommand(helloWorldCmdItem)

            ' Create another command site to hook the command to the right-click menu for Items.
            Dim itemContextCmdSite As New CommandSite("HelloWorldCommand.ItemContextMenu ", "Visor") With {
             .Location = CommandSiteLocation.ItemContextMenu,
             .DeployAsPulldownMenu = False
            }
            itemContextCmdSite.AddCommand(helloWorldCmdItem)

            ' Now the custom command is available in 3 places.

            'Gather the sites in a List.
            Dim sites As New List(Of CommandSite) From {
                toolbarCmdSite,
                fileContextCmdSite,
                itemContextCmdSite
            }

            ' Return the list of CommandSites.
            Return sites
        End Function


        ''' <summary>
        ''' This function tells Vault Explorer what custom tabs this extension provides.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <returns>A collection of DetailTabs, each object represents a custom tab.</returns>
        Public Function DetailTabs() As IEnumerable(Of DetailPaneTab) Implements IExplorerExtension.DetailTabs
            ' Create a DetailPaneTab list to return from method
            Dim fileTabs As New List(Of DetailPaneTab)()

            ' Return tabs
            Return fileTabs
        End Function

        ''' <summary>
        ''' This function is called after the user logs in to the Vault Server.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <param name="application">Provides information about the running application.</param>
        Public Sub OnLogOn(application As IApplication) Implements IExplorerExtension.OnLogOn

        End Sub

        ''' <summary>
        ''' This function is called after the user is logged out of the Vault Server.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <param name="application">Provides information about the running application.</param>
        Public Sub OnLogOff(application As IApplication) Implements IExplorerExtension.OnLogOff

        End Sub

        ''' <summary>
        ''' This function is called before the application is closed.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <param name="application">Provides information about the running application.</param>
        Public Sub OnShutdown(application As IApplication) Implements IExplorerExtension.OnShutdown
            ' Although this function is empty for this project, it's still needs to be defined 
            ' because it's part of the IExtension interface.
        End Sub

        ''' <summary>
        ''' This function is called after the application starts up.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <param name="application">Provides information about the running application.</param>
        Public Sub OnStartup(application As IApplication) Implements IExplorerExtension.OnStartup
            ' Although this function is empty for this project, it's still needs to be defined 
            ' because it's part of the IExtension interface.
        End Sub

        ''' <summary>
        ''' This function tells Vault Exlorer which default commands should be hidden.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <returns>A collection of command names.</returns>
        Public Function HiddenCommands() As IEnumerable(Of String) Implements IExplorerExtension.HiddenCommands
            ' This extension does not hide any commands.
            Return Nothing
        End Function

        ''' <summary>
        ''' This function allows the extension to define special behavior for Custom Entity types.
        ''' Part of the IExtension interface.
        ''' </summary>
        ''' <returns>A collection of CustomEntityHandler objects.  Each object defines special behavior
        ''' for a specific Custom Entity type.</returns>
        Public Function CustomEntityHandlers() As IEnumerable(Of CustomEntityHandler) Implements IExplorerExtension.CustomEntityHandlers
            ' This extension does not provide special Custom Entity behavior.
            Return Nothing
        End Function

#End Region


        ''' <summary>
        ''' This is the function that is called whenever the custom command is executed.
        ''' </summary>
        ''' <param name="s">The sender object.  Usually not used.</param>
        ''' <param name="e">The event args.  Provides additional information about the environment.</param>
        Private Sub HelloWorldCommandHandler(s As Object, e As CommandItemEventArgs)



            Try
                ' The Context part of the event args tells us information about what is selected.
                ' Run some checks to make sure that the selection is valid.

                If e.Context.CurrentSelectionSet.Count() = 0 Then
                    MessageBox.Show("No hay nada seleccionado")
                ElseIf e.Context.CurrentSelectionSet.Count() > 1 Then
                    MessageBox.Show("Esta función no soporta selecciones múltiples")
                Else

                    Dim proc = Process.GetProcessesByName("Rundll32")
                    For i As Integer = 0 To proc.Count - 1
                        proc(i).CloseMainWindow()
                    Next i

                    ' We only have one item selected, which is the expected behavior
                    Dim selection As ISelection = e.Context.CurrentSelectionSet.First()
                    Dim mgr As WebServiceManager = e.Context.Application.Connection.WebServiceManager

                    ' Look of the File or Item object. How we do this depends on what is selected.
                    Dim selectedFile As File = Nothing
                    Dim selectedItem As Item = Nothing
                    Dim revisionDef As PropDef = Nothing
                    Dim revisionProps As PropInst() = Nothing
                    Dim revisionProp As PropInst = Nothing
                    Dim fileName As String = Nothing
                    Dim itemName As String = Nothing
                    Dim revisionNumber As String = Nothing
                    Dim rutaFija As String = "R:\DTECNIC\PLANOS\0_PNG\"
                    Dim dirTres As String = Nothing
                    Dim dirSiete As String = Nothing
                    Dim longRevision As Integer = Nothing
                    Dim rutaImagen As String = Nothing

                    If selection.TypeId = SelectionTypeId.File Then
                        ' our ISelection.Id is really a File.MasterId
                        selectedFile = mgr.DocumentService.GetLatestFileByMasterId(selection.Id)
                        ' Encontrar la definición de propiedad "Revision" para archivos FILE
                        revisionDef = mgr.PropertyService.FindPropertyDefinitionsBySystemNames("FILE", New String() {"Revision"}).First()
                        revisionProps = mgr.PropertyService.GetProperties("FILE", New Long() {selectedFile.Id}, New Long() {revisionDef.Id})
                        revisionProp = revisionProps.First()
                        ' Mostrar el valor de la propiedad "Revision"
                        fileName = selectedFile.Name

                        fileName = Left(fileName, Len(fileName) - 4)
                        revisionNumber = revisionProp.Val.ToString()
                        If revisionNumber = "" Then
                            revisionNumber = "10"
                        End If
                        dirTres = Left(fileName, 3) & "\"
                        dirSiete = Left(fileName, 7) & "\"
                        longRevision = Len(revisionNumber)
                        If longRevision = 1 Then
                            rutaImagen = rutaFija & dirTres & dirSiete & fileName & "_R0" & revisionNumber & ".png"
                        Else
                            rutaImagen = rutaFija & dirTres & dirSiete & fileName & "_R" & revisionNumber & ".png"
                        End If

                        CheckPicturePath(rutaImagen)

                    ElseIf selection.TypeId = SelectionTypeId.FileVersion Then
                        ' our ISelection.Id is really a File.Id
                        selectedFile = mgr.DocumentService.GetFileById(selection.Id)
                        ' Encontrar la definición de propiedad "Revision" para archivos FILE
                        revisionDef = mgr.PropertyService.FindPropertyDefinitionsBySystemNames("FILE", New String() {"Revision"}).First()
                        revisionProps = mgr.PropertyService.GetProperties("FILE", New Long() {selectedFile.Id}, New Long() {revisionDef.Id})
                        revisionProp = revisionProps.First()
                        fileName = selectedFile.Name
                        fileName = Left(fileName, Len(fileName) - 4)
                        revisionNumber = revisionProp.Val.ToString()

                        If revisionNumber = "" Then
                            revisionNumber = "10"
                        End If

                        dirTres = Left(fileName, 3) & "\"
                        dirSiete = Left(fileName, 7) & "\"
                        longRevision = Len(revisionNumber)
                        If longRevision = 1 Then
                            rutaImagen = rutaFija & dirTres & dirSiete & fileName & "_R0" & revisionNumber & ".png"
                        Else
                            rutaImagen = rutaFija & dirTres & dirSiete & fileName & "_R" & revisionNumber & ".png"
                        End If

                        CheckPicturePath(rutaImagen)

                    ElseIf selection.TypeId = SelectionTypeId.Item Then
                        ' our ISelection.Id is really a Item.MasterId
                        selectedItem = mgr.ItemService.GetLatestItemByItemNumber(selection.Label)

                        ' Encontrar la definición de propiedad "Revision" para modelo ITEM
                        revisionDef = mgr.PropertyService.FindPropertyDefinitionsBySystemNames("ITEM", New String() {"RevNumber"}).First()
                        revisionProps = mgr.PropertyService.GetProperties("ITEM", New Long() {selectedItem.Id}, New Long() {revisionDef.Id})
                        revisionProp = revisionProps.First()
                        itemName = selectedItem.ItemNum
                        revisionNumber = revisionProp.Val.ToString()
                        dirTres = Left(itemName, 3) & "\"
                        dirSiete = Left(itemName, 7) & "\"
                        longRevision = Len(revisionNumber)
                        If longRevision = 1 Then
                            rutaImagen = rutaFija & dirTres & dirSiete & itemName & "_R0" & revisionNumber & ".png"
                        Else
                            rutaImagen = rutaFija & dirTres & dirSiete & itemName & "_R" & revisionNumber & ".png"
                        End If

                        CheckPicturePath(rutaImagen)

                    ElseIf selection.TypeId = SelectionTypeId.Bom Then
                        selectedItem = mgr.ItemService.GetLatestItemByItemNumber(selection.Label)
                        ' Encontrar la definición de propiedad "Revision" para modelo ITEM
                        revisionDef = mgr.PropertyService.FindPropertyDefinitionsBySystemNames("ITEM", New String() {"RevNumber"}).First()
                        revisionProps = mgr.PropertyService.GetProperties("ITEM", New Long() {selectedItem.Id}, New Long() {revisionDef.Id})
                        revisionProp = revisionProps.First()
                        itemName = selectedItem.ItemNum
                        revisionNumber = revisionProp.Val.ToString()
                        dirTres = Left(itemName, 3) & "\"
                        dirSiete = Left(itemName, 7) & "\"
                        longRevision = Len(revisionNumber)
                        If longRevision = 1 Then
                            rutaImagen = rutaFija & dirTres & dirSiete & itemName & "_R0" & revisionNumber & ".png"
                        Else
                            rutaImagen = rutaFija & dirTres & dirSiete & itemName & "_R" & revisionNumber & ".png"
                        End If

                        CheckPicturePath(rutaImagen)

                    End If

                    'If selectedFile Is Nothing Then
                    '    MessageBox.Show("La selección no es un fichero.")
                    'End If

                End If
            Catch ex As Exception
                ' If something goes wrong, we don't want the exception to bubble up to Vault Explorer.
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Sub

        Private Sub CheckPicturePath(rutaImagen As String)
            ' Verificar si el archivo de imagen existe en rutaImagen
            If System.IO.File.Exists(rutaImagen) Then
                ' El archivo de imagen existe, ejecutar el visor de imágenes
                Dim proceso As New Process()
                proceso.StartInfo.FileName = "C:\Program Files (x86)\FastStone Image Viewer\FSViewer.exe"
                proceso.StartInfo.Arguments = rutaImagen
                proceso.Start()
            Else
                ' El archivo de imagen no existe, mostrar un mensaje o realizar otra acción
                MessageBox.Show("NO EXISTE IMAGEN DEL PLANO.", "IMAGEN NO ENCONTRADA", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

    End Class
End Namespace
