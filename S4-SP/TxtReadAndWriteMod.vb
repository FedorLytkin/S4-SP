Imports System
Imports System.IO
Module TxtReadAndWriteMod
    Sub txt_create()
        Dim w As StreamWriter = New StreamWriter("C:\Users\aidarhanov.n.VEZA-SPB\Desktop\test 1\append.txt")
        w.WriteLine("ArtID  Display Name           Driver Type   Link Date")
        w.WriteLine("============ ====================== ============= ======================")
        w.WriteLine("{0,-12} {1}")
        w.Close()
    End Sub
End Module
