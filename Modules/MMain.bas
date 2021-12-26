Attribute VB_Name = "MMain"
Option Explicit

Sub Main()
    FMain.Show
End Sub

Public Function ImagePNM(aPFN As String) As ImagePNM
    Set ImagePNM = New ImagePNM: ImagePNM.New_ aPFN
End Function
