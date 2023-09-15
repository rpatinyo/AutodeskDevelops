'=====================================================================
'
'  This file is part of the Autodesk Vault API Code Samples.
'
'  Copyright (C) Autodesk Inc.  All rights reserved.
'
'THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
'KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'PARTICULAR PURPOSE.
'=====================================================================

Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Imports Autodesk.DataManagement.Client.Framework.Forms

Namespace HelloWorld
    Partial Public Class MyCustomTabControl
        Inherits UserControl

        Private mColor As Color

        Public Sub New()
            InitializeComponent()
            mColor = mPropertyGrid.LineColor

            SetTheme(Library.CurrentTheme)
            AddHandler Library.ThemeChanged, AddressOf ThemeChangedHandler
        End Sub

        Public Sub SetSelectedObject(ByVal o As Object)
            mPropertyGrid.SelectedObject = o
        End Sub
        Friend WithEvents mPropertyGrid As System.Windows.Forms.PropertyGrid

        Private Sub ThemeChangedHandler(sender As Object, theme As Library.UITheme)
            SetTheme(theme)
        End Sub

        Private Sub SetTheme(theme As Library.UITheme)
            ' The control Is already responding to theme change,
            ' but only property grid's column header is not.
            ' Adjust its appearence explicitly so it looks better in Dark theme.
            Select Case theme
                Case Library.UITheme.Dark
                    mPropertyGrid.CategoryForeColor = Color.White
                    mPropertyGrid.LineColor = Color.CadetBlue
                Case Library.UITheme.Light, Library.UITheme.Classic
                    mPropertyGrid.CategoryForeColor = Color.Black
                    mPropertyGrid.LineColor = mColor
            End Select
        End Sub

        Private Sub InitializeComponent()
            Me.mPropertyGrid = New System.Windows.Forms.PropertyGrid()
            Me.SuspendLayout()
            '
            'mPropertyGrid
            '
            Me.mPropertyGrid.Location = New System.Drawing.Point(4, 4)
            Me.mPropertyGrid.Name = "mPropertyGrid"
            Me.mPropertyGrid.Size = New System.Drawing.Size(231, 230)
            Me.mPropertyGrid.TabIndex = 0
            '
            'MyCustomTabControl
            '
            Me.Controls.Add(Me.mPropertyGrid)
            Me.Name = "MyCustomTabControl"
            Me.Size = New System.Drawing.Size(238, 237)
            Me.ResumeLayout(False)

        End Sub
    End Class
End Namespace
