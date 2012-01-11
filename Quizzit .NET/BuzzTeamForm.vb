Public Class BuzzTeamForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TeamLabel As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.TeamLabel = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'TeamLabel
        '
        Me.TeamLabel.BackColor = System.Drawing.Color.Black
        Me.TeamLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TeamLabel.Font = New System.Drawing.Font("Copperplate Gothic Bold", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TeamLabel.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.TeamLabel.Location = New System.Drawing.Point(24, 16)
        Me.TeamLabel.Name = "TeamLabel"
        Me.TeamLabel.Size = New System.Drawing.Size(248, 65)
        Me.TeamLabel.TabIndex = 0
        Me.TeamLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Timer1
        '
        Me.Timer1.Interval = 1500
        '
        'BuzzTeamForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(300, 100)
        Me.Controls.Add(Me.TeamLabel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "BuzzTeamForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BuzzTeamForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub BuzzTeamForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TeamLabel.Text = "Team " + TeamBtn
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        Me.Hide()
        TeamLabel.Text = ""
    End Sub

    Private Sub BuzzTeamForm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        PlaySnd(ResourcePath + "\Tones\glass.wav")
    End Sub
End Class
