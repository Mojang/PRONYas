Public Class clsMain

    Public Shared gfrmMain As Form1



    Shared Sub Main()

        clsMain.gfrmMain = New Form1

        clsMain.gfrmMain.ShowDialog()

    End Sub


End Class
