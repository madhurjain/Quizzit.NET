Public Class qzMain
    Inherits System.Windows.Forms.Form
    Dim StatusBarEnable As Boolean


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        Application.EnableVisualStyles()
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
    Friend WithEvents qzMainMenu As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents StatTimer As System.Windows.Forms.Timer
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel4 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents AxShockwaveFlashMain As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(qzMain))
        Me.qzMainMenu = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.MenuItem15 = New System.Windows.Forms.MenuItem()
        Me.MenuItem16 = New System.Windows.Forms.MenuItem()
        Me.MenuItem17 = New System.Windows.Forms.MenuItem()
        Me.MenuItem13 = New System.Windows.Forms.MenuItem()
        Me.MenuItem20 = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.MenuItem21 = New System.Windows.Forms.MenuItem()
        Me.MenuItem22 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.MenuItem14 = New System.Windows.Forms.MenuItem()
        Me.MenuItem18 = New System.Windows.Forms.MenuItem()
        Me.MenuItem19 = New System.Windows.Forms.MenuItem()
        Me.StatusBar1 = New System.Windows.Forms.StatusBar()
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarPanel3 = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarPanel4 = New System.Windows.Forms.StatusBarPanel()
        Me.StatTimer = New System.Windows.Forms.Timer(Me.components)
        Me.AxShockwaveFlashMain = New AxShockwaveFlashObjects.AxShockwaveFlash()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxShockwaveFlashMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'qzMainMenu
        '
        Me.qzMainMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4, Me.MenuItem5})
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem13, Me.MenuItem20, Me.MenuItem11, Me.MenuItem21, Me.MenuItem22})
        Me.MenuItem4.Text = "&Setup"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3, Me.MenuItem9, Me.MenuItem15, Me.MenuItem16, Me.MenuItem17})
        Me.MenuItem1.Text = "Select R&ound"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.RadioCheck = True
        Me.MenuItem2.Shortcut = System.Windows.Forms.Shortcut.CtrlF1
        Me.MenuItem2.Text = "SemiFinal I"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.RadioCheck = True
        Me.MenuItem3.Shortcut = System.Windows.Forms.Shortcut.CtrlF2
        Me.MenuItem3.Text = "SemiFinal II"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 2
        Me.MenuItem9.RadioCheck = True
        Me.MenuItem9.Shortcut = System.Windows.Forms.Shortcut.CtrlF3
        Me.MenuItem9.Text = "SemiFinal III"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 3
        Me.MenuItem15.RadioCheck = True
        Me.MenuItem15.Shortcut = System.Windows.Forms.Shortcut.CtrlF4
        Me.MenuItem15.Text = "SemiFinal IV"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 4
        Me.MenuItem16.RadioCheck = True
        Me.MenuItem16.Shortcut = System.Windows.Forms.Shortcut.CtrlF5
        Me.MenuItem16.Text = "Final"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 5
        Me.MenuItem17.RadioCheck = True
        Me.MenuItem17.Shortcut = System.Windows.Forms.Shortcut.CtrlF6
        Me.MenuItem17.Text = "College"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 1
        Me.MenuItem13.Text = "H/W Test"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 2
        Me.MenuItem20.Text = "Panic ScoreSet"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 3
        Me.MenuItem11.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.MenuItem11.Text = "Score Card"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 4
        Me.MenuItem21.Text = "Clear Score"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 5
        Me.MenuItem22.Text = "Status Bar"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem7, Me.MenuItem8, Me.MenuItem10, Me.MenuItem12, Me.MenuItem14, Me.MenuItem18, Me.MenuItem19})
        Me.MenuItem5.Text = "&Rounds"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Shortcut = System.Windows.Forms.Shortcut.CtrlG
        Me.MenuItem6.Text = "General"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Shortcut = System.Windows.Forms.Shortcut.CtrlH
        Me.MenuItem7.Text = "Hangman"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Shortcut = System.Windows.Forms.Shortcut.CtrlV
        Me.MenuItem8.Text = "Video"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 3
        Me.MenuItem10.Shortcut = System.Windows.Forms.Shortcut.CtrlA
        Me.MenuItem10.Text = "Audio"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 4
        Me.MenuItem12.Shortcut = System.Windows.Forms.Shortcut.CtrlW
        Me.MenuItem12.Text = "Windfall"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 5
        Me.MenuItem14.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.MenuItem14.Text = "RapidFire"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 6
        Me.MenuItem18.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.MenuItem18.Text = "CrossWordz"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 7
        Me.MenuItem19.Shortcut = System.Windows.Forms.Shortcut.CtrlI
        Me.MenuItem19.Text = "Image"
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 438)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2, Me.StatusBarPanel3, Me.StatusBarPanel4})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(816, 22)
        Me.StatusBar1.SizingGrip = False
        Me.StatusBar1.TabIndex = 1
        Me.StatusBar1.Visible = False
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel1.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.StatusBarPanel1.MinWidth = 200
        Me.StatusBarPanel1.Name = "StatusBarPanel1"
        Me.StatusBarPanel1.Width = 200
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel2.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.StatusBarPanel2.MinWidth = 200
        Me.StatusBarPanel2.Name = "StatusBarPanel2"
        Me.StatusBarPanel2.Width = 200
        '
        'StatusBarPanel3
        '
        Me.StatusBarPanel3.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel3.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.StatusBarPanel3.MinWidth = 200
        Me.StatusBarPanel3.Name = "StatusBarPanel3"
        Me.StatusBarPanel3.Width = 200
        '
        'StatusBarPanel4
        '
        Me.StatusBarPanel4.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel4.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.StatusBarPanel4.MinWidth = 200
        Me.StatusBarPanel4.Name = "StatusBarPanel4"
        Me.StatusBarPanel4.Width = 200
        '
        'StatTimer
        '
        Me.StatTimer.Interval = 500
        '
        'AxShockwaveFlashMain
        '
        Me.AxShockwaveFlashMain.Enabled = True
        Me.AxShockwaveFlashMain.Location = New System.Drawing.Point(26, 38)
        Me.AxShockwaveFlashMain.Name = "AxShockwaveFlashMain"
        Me.AxShockwaveFlashMain.OcxState = CType(resources.GetObject("AxShockwaveFlashMain.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlashMain.Size = New System.Drawing.Size(766, 359)
        Me.AxShockwaveFlashMain.TabIndex = 20
        '
        'qzMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(816, 460)
        Me.Controls.Add(Me.AxShockwaveFlashMain)
        Me.Controls.Add(Me.StatusBar1)
        Me.ForeColor = System.Drawing.Color.White
        Me.IsMdiContainer = True
        Me.Menu = Me.qzMainMenu
        Me.MinimizeBox = False
        Me.Name = "qzMain"
        Me.Text = "Quizzit .NET"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxShockwaveFlashMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub qzMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '=========================================
        'To Change The Background Color Of MDI Parent
        '=========================================
        For Each ctl As Control In Me.Controls
            Try
                If ctl.GetType Is GetType(Windows.Forms.MdiClient) Then
                    ctl.BackColor = Color.Black
                    Exit For
                End If
            Catch exc As InvalidCastException
                'Catch and ignore the error if casting failed
            End Try
        Next
        StatusBarEnable = False
        AxShockwaveFlashMain.Show()
        AxShockwaveFlashMain.Width = Me.Width
        AxShockwaveFlashMain.Height = Me.Height
        AxShockwaveFlashMain.Movie = ResourcePath + "\Extras\Brain.swf"
        AxShockwaveFlashMain.Play()
    End Sub
    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndgen As New rndGeneralForm
            rndgen.MdiParent = Me
            rndgen.WindowState = FormWindowState.Maximized
            rndgen.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !"
        End If
    End Sub
    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndhan As New rndHangManForm
            rndhan.MdiParent = Me
            rndhan.WindowState = FormWindowState.Maximized
            rndhan.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !"
        End If
    End Sub
    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndvid As New rndVideoForm
            rndvid.MdiParent = Me
            rndvid.WindowState = FormWindowState.Maximized
            rndvid.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !"
        End If
    End Sub
    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndaud As New rndAudioForm
            rndaud.MdiParent = Me
            rndaud.WindowState = FormWindowState.Maximized
            rndaud.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !"
        End If
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndrap As New rndRapidFireForm
            rndrap.MdiParent = Me
            rndrap.WindowState = FormWindowState.Maximized
            rndrap.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !"
        End If
    End Sub
    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndWin As New rndWindfallForm
            rndWin.MdiParent = Me
            rndWin.WindowState = FormWindowState.Maximized
            rndWin.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !"
        End If
    End Sub
    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndCRW As New rndCrossWord
            rndCRW.MdiParent = Me
            rndCRW.WindowState = FormWindowState.Maximized
            rndCRW.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !"
        End If
    End Sub
    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem19.Click
        ClearStats()
        If Me.ActiveMdiChild Is Nothing Then
            Dim rndIMG As New rndImageForm
            rndIMG.MdiParent = Me
            rndIMG.WindowState = FormWindowState.Maximized
            rndIMG.Show()
        Else
            Stats2 = Me.ActiveMdiChild.Text + " Window Already Open !" + Chr(13)
        End If
    End Sub
    Private Sub MenuItem11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        Dim ScrFrm As New ScoreForm
        ScrFrm.WindowState = FormWindowState.Maximized
        ScrFrm.Show()
    End Sub
    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        Dim hwForm As New HWTestForm
        hwForm.Show()
    End Sub
    Private Sub MenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem20.Click
        Dim PanScoreFrm As New PanicScoresForm
        PanScoreFrm.Show()
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        MainRound = "SEMIFINAL1"
    End Sub
    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        MainRound = "SEMIFINAL2"
    End Sub
    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        MainRound = "SEMIFINAL3"
    End Sub
    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click
        MainRound = "SEMIFINAL4"
    End Sub
    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click
        MainRound = "FINAL"
    End Sub
    Private Sub MenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem17.Click
        MainRound = "COLLEGE"
    End Sub
    Private Sub StatTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatTimer.Tick
        StatusBarPanel1.Text = "Team " + TeamBtn
        StatusBarPanel2.Text = Stats2
        StatusBarPanel3.Text = Stats3
        StatusBarPanel4.Text = A_Score & " ," & B_Score & " ," & C_Score & " ," & D_Score
    End Sub
    Private Sub ClearStats()
        AxShockwaveFlashMain.Hide()
        StatusBarPanel1.Text = ""
        Stats2 = ""
        Stats3 = ""
        StatusBarPanel4.Text = ""
    End Sub
    Private Sub qzMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        StatTimer.Stop()
    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem21.Click

        Try
            OpenConnection()
            OpenScoreSet("Update Scores set A=0,B=0,C=0,D=0")
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error In Closing Connection/Recordset")
        End Try
        Try
            CloseScoreSet()
            CloseConnection()
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error In Closing Connection/Recordset")
        End Try
    End Sub

    Private Sub MenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem22.Click
        If StatusBarEnable = True Then
            StatusBar1.Visible = False
            StatTimer.Stop()
            StatusBarEnable = False
        Else
            StatusBar1.Visible = True
            StatTimer.Start()
            StatusBarEnable = True
        End If
    End Sub
End Class
