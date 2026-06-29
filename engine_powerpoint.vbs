' ============================================================================
' engine_powerpoint.vbs  --  Converty PNG->PowerPoint assembly engine
'
' Builds a .pptx where each page image becomes one full-bleed slide.
' Runs under cscript.exe for the same reliability reason as the Publisher engine.
'
' Usage:  cscript //B engine_powerpoint.vbs <manifest> <logfile>
'   manifest : UTF-16 text, one job per line:
'                <pptxOutPath>|<pngDir>|<slideWpts>|<slideHpts>|<pageCount>
'              pngDir must contain page-1.png .. page-<pageCount>.png
'   logfile  : UTF-16 text; status lines:
'                OK|<pptxPath>          deck written
'                ERR|<pptxPath>|<reason>
'                DONE_FILE              after every job (progress)
'                FATAL|<reason>         PowerPoint could not start
'                ALL_DONE
' ============================================================================
Option Explicit

Const ForReading = 1
Const ForAppending = 8
Const TristateTrue = -1
Const msoFalse = 0
Const msoTrue = -1
Const ppLayoutBlank = 12
Const ppSaveAsOpenXMLPresentation = 24

Dim fso, args, manifestPath, logPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments
If args.Count < 2 Then WScript.Quit 2
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

Dim ppt
On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Or ppt Is Nothing Then
  Log "FATAL|Cannot start Microsoft PowerPoint: " & Err.Description
  WScript.Quit 1
End If
On Error Goto 0

Dim ts, line, parts, pptxPath, pngDir, wPts, hPts, pageCount
Dim pres, i, png, slide

Set ts = fso.OpenTextFile(manifestPath, ForReading, False, TristateTrue)
Do Until ts.AtEndOfStream
  line = ts.ReadLine
  If Len(Trim(line)) > 0 Then
    parts = Split(line, "|")
    If UBound(parts) >= 4 Then
      pptxPath  = parts(0)
      pngDir    = parts(1)
      wPts      = CDbl(parts(2))
      hPts      = CDbl(parts(3))
      pageCount = CLng(parts(4))

      On Error Resume Next
      Err.Clear
      Set pres = ppt.Presentations.Add(msoFalse)   ' no window
      If Err.Number <> 0 Then
        Log "ERR|" & pptxPath & "|Could not create presentation: " & Err.Description
      Else
        pres.PageSetup.SlideWidth = wPts
        pres.PageSetup.SlideHeight = hPts
        For i = 1 To pageCount
          png = fso.BuildPath(pngDir, "page-" & i & ".png")
          If fso.FileExists(png) Then
            Set slide = pres.Slides.Add(i, ppLayoutBlank)
            slide.Shapes.AddPicture png, msoFalse, msoTrue, 0, 0, wPts, hPts
          End If
        Next
        ' Remove any stale output so PowerPoint never shows an overwrite prompt.
        If fso.FileExists(pptxPath) Then fso.DeleteFile pptxPath, True
        Err.Clear
        pres.SaveAs pptxPath, ppSaveAsOpenXMLPresentation
        If Err.Number <> 0 Then
          Log "ERR|" & pptxPath & "|Save failed: " & Err.Description
        Else
          Log "OK|" & pptxPath
        End If
        pres.Close
      End If
      On Error Goto 0

      Log "DONE_FILE"
    End If
  End If
Loop
ts.Close

On Error Resume Next
ppt.Quit
On Error Goto 0
Log "ALL_DONE"
WScript.Quit 0
