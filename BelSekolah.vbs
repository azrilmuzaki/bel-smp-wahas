' Bel Sekolah Otomatis - SMP Wahid Hasyim Kepohbaru
' VBScript untuk Windows 7. Satu folder dengan jadwal.txt dan file mp3.
Option Explicit

Dim fso, folder, player, ts, lines(), cnt, ln, p, d, nowSec, st, sp, f, k
Dim days, curWin, curStop, doneKeys, today, lastDate, i

Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
Set player = CreateObject("WMPlayer.OCX")
player.settings.volume = 100
curWin = ""
curStop = 0
doneKeys = ""
lastDate = ""

Function ToSec(t)
    Dim a
    a = Split(Trim(t), ":")
    ToSec = CLng(a(0)) * 3600 + CLng(a(1)) * 60
End Function

Sub PlayFull(file)
    Dim n
    player.URL = file
    player.controls.play
    WScript.Sleep 1500
    n = 0
    Do While (player.playState = 3 Or player.playState = 6 Or player.playState = 9) And n < 3600
        WScript.Sleep 500
        n = n + 1
    Loop
    player.controls.stop
End Sub

Sub StartAt(file, offset)
    Dim n
    player.URL = file
    player.controls.play
    n = 0
    Do While player.playState <> 3 And n < 40
        WScript.Sleep 250
        n = n + 1
    Loop
    If offset > 0 Then player.controls.currentPosition = offset
End Sub

Do
    On Error Resume Next
    today = CStr(Date)
    If today <> lastDate Then
        doneKeys = ""
        lastDate = today
    End If

    d = Weekday(Now, vbMonday)   ' 1=Senin ... 6=Sabtu, 7=Ahad
    nowSec = Hour(Now) * 3600 + Minute(Now) * 60 + Second(Now)

    ' hentikan pemutaran berdurasi (waqiah) jika waktunya habis / file selesai
    If curWin <> "" Then
        If nowSec >= curStop Or player.playState = 1 Or player.playState = 8 Then
            player.controls.stop
            curWin = ""
        End If
    End If

    ' baca jadwal (dibaca ulang tiap detik, jadi edit langsung berlaku)
    cnt = 0
    ReDim lines(200)
    If fso.FileExists(folder & "\jadwal.txt") Then
        Set ts = fso.OpenTextFile(folder & "\jadwal.txt", 1)
        Do While Not ts.AtEndOfStream And cnt < 200
            ln = Trim(ts.ReadLine)
            If Len(ln) > 0 And Left(ln, 1) <> "#" Then
                lines(cnt) = ln
                cnt = cnt + 1
            End If
        Loop
        ts.Close
    End If

    For i = 0 To cnt - 1
        p = Split(lines(i), ",")
        If UBound(p) >= 2 Then
            days = Trim(p(0))
            If days = "*" Or InStr(days, CStr(d)) > 0 Then
                st = ToSec(p(1))
                f = folder & "\" & Trim(p(2))
                k = "[" & today & "|" & lines(i) & "]"
                If InStr(doneKeys, k) = 0 And fso.FileExists(f) Then
                    If UBound(p) >= 3 Then
                        ' pemutaran berdurasi: mulai dari posisi sesuai jam
                        sp = ToSec(p(3))
                        If curWin = "" And nowSec >= st And nowSec < sp Then
                            doneKeys = doneKeys & k
                            StartAt f, nowSec - st
                            curWin = k
                            curStop = sp
                        End If
                    Else
                        ' bel biasa
                        If curWin = "" And nowSec >= st And nowSec < st + 30 Then
                            doneKeys = doneKeys & k
                            PlayFull f
                        End If
                    End If
                End If
            End If
        End If
    Next

    WScript.Sleep 1000
Loop
