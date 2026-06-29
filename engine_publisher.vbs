' ============================================================================
' engine_publisher.vbs  --  Converty Publisher->PDF conversion engine
'
' Runs under cscript.exe (NOT PowerShell -- Publisher COM deadlocks under
' PowerShell's STA because it does not pump the message loop the way Office's
' synchronous file serialization requires).
'
' Usage:  cscript //B engine_publisher.vbs <manifest> <logfile>
'   manifest : UTF-16 text, one job per line:  <pubPath>|<pdfOutPath>
'   logfile  : UTF-16 text; engine appends status lines the GUI tails:
'                OK|<pubPath>|<pdfPath>      file converted
'                ERR|<pubPath>|<reason>      file failed (non-fatal, continues)
'                DONE_FILE                   emitted after every job (progress)
'                FATAL|<reason>              Publisher could not start
'                ALL_DONE                    whole manifest finished
' ============================================================================
Option Explicit

Const ForReading = 1
Const ForAppending = 8
Const TristateTrue = -1            ' open as Unicode (UTF-16)
Const pbFixedFormatTypePDF = 2

Dim fso, args, manifestPath, logPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments

If args.Count < 2 Then
  WScript.Quit 2
End If
manifestPath = args(0)
logPath = args(1)

Sub Log(s)
  On Error Resume Next
  Dim f
  Set f = fso.OpenTextFile(logPath, ForAppending, True, TristateTrue)
  f.WriteLine s
  f.Close
  On Error Goto 0
End Sub

Dim pub
On Error Resume Next
Set pub = CreateObject("Publisher.Application")
If Err.Number <> 0 Or pub Is Nothing Then
  Log "FATAL|Cannot start Microsoft Publisher: " & Err.Description
  WScript.Quit 1
End If
On Error Goto 0

Dim ts, line, parts, pubPath, pdfPath, doc
Set ts = fso.OpenTextFile(manifestPath, ForReading, False, TristateTrue)

Do Until ts.AtEndOfStream
  line = ts.ReadLine
  If Len(Trim(line)) > 0 Then
    parts = Split(line, "|")
    If UBound(parts) >= 1 Then
      pubPath = parts(0)
      pdfPath = parts(1)

      On Error Resume Next
      Err.Clear
      Set doc = Nothing
      Set doc = pub.Open(pubPath)
      If Err.Number <> 0 Or doc Is Nothing Then
        Log "ERR|" & pubPath & "|Could not open: " & Err.Description
      Else
        ' Remove any stale output so Publisher never shows an overwrite prompt.
        If fso.FileExists(pdfPath) Then fso.DeleteFile pdfPath, True
        Err.Clear
        doc.ExportAsFixedFormat pbFixedFormatTypePDF, pdfPath
        If Err.Number <> 0 Then
          Log "ERR|" & pubPath & "|PDF export failed: " & Err.Description
        Else
          Log "OK|" & pubPath & "|" & pdfPath
        End If
        ' (Export does not modify the publication, so Close does not prompt.)
        doc.Close
      End If
      On Error Goto 0

      Log "DONE_FILE"
    End If
  End If
Loop
ts.Close

On Error Resume Next
pub.Quit
On Error Goto 0
Log "ALL_DONE"
WScript.Quit 0
