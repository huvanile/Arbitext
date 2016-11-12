Imports Arbitext.RegistryHelpers

Public Class PushbulletHelpers


    Public Shared Sub sendPushbulletNote(noteTitle As String, noteBody As String)
        '    If thisaddin.notifyViaPBOK = True Then
        '        Dim resp As String
        '        Dim apiKey As String : apiKey = thisaddin.pbapikey
        '        Dim device As String : device = thisaddin.pbDeviceID
        '        If device = "" And IsNumeric(device) Then
        '            resp = ShellRun("pythonw -c ""from pushbullet import Pushbullet ; pb=Pushbullet('" & apiKey & "') ; push = pb.push_note('" & noteTitle & "','" & noteBody & "') """)
        '        Else
        '            resp = ShellRun("pythonw -c ""from pushbullet import Pushbullet ; pb=Pushbullet('" & apiKey & "') ; theD = pb.devices[" & device & "] ; push = theD.push_note('" & noteTitle & "','" & noteBody & "') """)
        '        End If
        '        Debug.Print resp
        'End If
        '    'OLD VERSION, REQUIRES CURL TO BE INSTALLED
        '    '        Dim wsh As Object
        '    '        Set wsh = VBA.CreateObject("WScript.Shell")
        '    '        Dim waitOnReturn As Boolean: waitOnReturn = True
        '    '        Dim windowStyle As Integer: windowStyle = 1
        '    '        wsh.Run "cmd.exe /S /C curl.exe https://api.pushbullet.com/v2/pushes -X POST -u " & apiKey & ": --header ""Content-Type: application/json"" --data-binary ""{\""device_iden\"":\""" & device & "\"", \""type\"":\""note\"", \""title\"":\""" & noteTitle & "\"", \""body\"": \""" & noteBody & "\""}""", windowStyle, waitOnReturn
    End Sub

End Class
