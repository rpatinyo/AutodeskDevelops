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

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MyCustomTabControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.mPropertyGrid = New System.Windows.Forms.PropertyGrid
        Me.SuspendLayout()
        '
        'mPropertyGrid
        '
        Me.mPropertyGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.mPropertyGrid.Location = New System.Drawing.Point(0, 0)
        Me.mPropertyGrid.Name = "mPropertyGrid"
        Me.mPropertyGrid.Size = New System.Drawing.Size(336, 150)
        Me.mPropertyGrid.TabIndex = 0
        '
        'MyCustomTabControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.mPropertyGrid)
        Me.Name = "MyCustomTabControl"
        Me.Size = New System.Drawing.Size(336, 150)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents mPropertyGrid As System.Windows.Forms.PropertyGrid

End Class
