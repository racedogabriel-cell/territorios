Option Explicit
Dim fso, shell, base, ps, cmd, checkCmd, rc
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

base = fso.GetParentFolderName(WScript.ScriptFullName)
ps = shell.ExpandEnvironmentStrings("%WINDIR%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"

' Si el servidor ya esta arriba, no intenta relanzar.
checkCmd = Chr(34) & ps & Chr(34) & " -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & _
           "try { Invoke-RestMethod -Uri 'http://localhost:8080/api/ping' -TimeoutSec 2 | Out-Null; exit 0 } catch { exit 1 }" & _
           Chr(34)
rc = shell.Run(checkCmd, 0, True)

If rc <> 0 Then
  cmd = Chr(34) & ps & Chr(34) & " -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & base & "\servir.ps1" & Chr(34) & " -Puerto 8080"
  shell.Run cmd, 0, False
  WaitForServer 8080, 40, 250
Else
  shell.Run "http://localhost:8080/mapa.html?v=" & CStr(Timer), 1, False
End If

Function IsServerUp(port)
  On Error Resume Next
  Dim http, url
  Set http = CreateObject("MSXML2.XMLHTTP")
  url = "http://localhost:" & CStr(port) & "/api/ping"
  http.Open "GET", url, False
  http.Send
  If Err.Number <> 0 Then
    Err.Clear
    IsServerUp = False
  Else
    IsServerUp = (http.Status = 200)
  End If
End Function

Sub WaitForServer(port, maxTry, sleepMs)
  Dim i
  For i = 1 To maxTry
    If IsServerUp(port) Then Exit For
    WScript.Sleep sleepMs
  Next
End Sub
