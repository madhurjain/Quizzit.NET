Module SndModule
    Private Const SND_APPLICATION As Integer = &H80         '  look for application specific association
    Private Const SND_ALIAS As Integer = &H10000     '  name is a WIN.INI [sounds] entry
    Private Const SND_ALIAS_ID As Integer = &H110000    '  name is a WIN.INI [sounds] entry identifier
    Private Const SND_ASYNC As Integer = &H1         '  play asynchronously
    Private Const SND_FILENAME As Integer = &H20000     '  name is a file name
    Private Const SND_LOOP As Integer = &H8         '  loop the sound until next sndPlaySound
    Private Const SND_MEMORY As Integer = &H4         '  lpszSoundName points to a memory file
    Private Const SND_NODEFAULT As Integer = &H2         '  silence not default, if sound not found
    Private Const SND_NOSTOP As Integer = &H10        '  don't stop any currently playing sound
    Private Const SND_NOWAIT As Integer = &H2000      '  don't wait if the driver is busy
    Private Const SND_PURGE As Integer = &H40               '  purge non-static events for task
    Private Const SND_RESOURCE As Integer = &H40004     '  name is a resource name or atom
    Private Const SND_SYNC As Integer = &H0         '  play synchronously (default)
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
    Public Sub PlaySnd(ByVal File As String)
        Application.DoEvents()
        PlaySound(File, 0&, SND_FILENAME Or SND_ASYNC)
    End Sub


End Module
