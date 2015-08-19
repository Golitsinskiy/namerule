Option Explicit On
Imports System.IO

Public Class Logging

    Public Sub WriteLog(Message As String, mFileName As String)

        Dim FileName As String = Path.GetTempPath() & mFileName
        Dim objStreamWriter As StreamWriter

        If File.Exists(FileName) = False Then
            objStreamWriter = File.CreateText(FileName)
            objStreamWriter.Flush()
            objStreamWriter.Close()
        End If

        objStreamWriter = File.AppendText(FileName)
        objStreamWriter.WriteLine(DateTime.Now.ToShortDateString & " " & DateTime.Now.ToLongTimeString & " " & Message)
        objStreamWriter.Close()

    End Sub

End Class
