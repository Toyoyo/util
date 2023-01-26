Attribute VB_Name = "Module2"
Sub main()
    Dim FrmMain As New MainForm
    Dim nPosX As Integer
    Dim nPosY As Integer
    
    FrmMain.left = Screen.Width - FrmMain.Width
    FrmMain.top = 0
    FrmMain.Show
End Sub
