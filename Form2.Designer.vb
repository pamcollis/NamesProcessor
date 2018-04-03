<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.lvProcessedNames = New System.Windows.Forms.ListView()
        Me.cName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cDateTime = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cStatus = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cDescription = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'lvProcessedNames
        '
        Me.lvProcessedNames.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.cName, Me.cDateTime, Me.cStatus, Me.cDescription})
        Me.lvProcessedNames.Location = New System.Drawing.Point(13, 23)
        Me.lvProcessedNames.Name = "lvProcessedNames"
        Me.lvProcessedNames.Size = New System.Drawing.Size(509, 227)
        Me.lvProcessedNames.TabIndex = 0
        Me.lvProcessedNames.UseCompatibleStateImageBehavior = False
        Me.lvProcessedNames.View = System.Windows.Forms.View.Details
        '
        'cName
        '
        Me.cName.Text = "Client Name"
        Me.cName.Width = 200
        '
        'cDateTime
        '
        Me.cDateTime.Text = "Date/Time"
        Me.cDateTime.Width = 100
        '
        'cStatus
        '
        Me.cStatus.Text = "Status"
        Me.cStatus.Width = 100
        '
        'cDescription
        '
        Me.cDescription.Text = "Description"
        Me.cDescription.Width = 300
        '
        'fMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(534, 262)
        Me.Controls.Add(Me.lvProcessedNames)
        Me.Name = "fMain"
        Me.Text = "Names Processor"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lvProcessedNames As System.Windows.Forms.ListView
    Friend WithEvents cName As System.Windows.Forms.ColumnHeader
    Friend WithEvents cDateTime As System.Windows.Forms.ColumnHeader
    Friend WithEvents cStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents cDescription As System.Windows.Forms.ColumnHeader
End Class
