#If VBA7 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Sub Extract_pdfs()

Dim wb As Workbook
Dim sh As Worksheet

Dim PrintRange As Range
Set PrintRange = Range("A1:O120")

Set wb = ThisWorkbook

For Each sh In wb.Worksheets
    If sh.Name Like "Sheet *" Then

        sh.Select
        
        
        With ActiveSheet.PageSetup
            .Zoom = False
            .PrintArea = PrintRange.Address
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .Orientation = xlLandscape
            .LeftMargin = Application.InchesToPoints(0.15)
            .RightMargin = Application.InchesToPoints(0.15)
            .TopMargin = Application.InchesToPoints(0.15)
            .BottomMargin = Application.InchesToPoints(0.15)
        End With
    
        namestring = "\" & sh.Range("R3").Value & ".pdf"
        pdf_name = sh.Name & ".pdf"
    
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=ThisWorkbook.Path & namestring, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    End If
Next

End Sub


Sub Send_Email_Using_VBA()
'Now with embedded report images
'don't forget to add Microsoft Object Library in Tools > References

Dim wb As Workbook
Dim sh As Worksheet

Set wb = ThisWorkbook

Dim Email_Subject, Email_Send_From, Email_Send_To, Email_Cc, Email_Body, Email_Attach As String
Dim htmlTemp, body1, body2 As String

Dim colAttach As Outlook.attachments
Dim oAttach As Outlook.attachment

Dim olkPA As Outlook.PropertyAccessor

Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

For Each sh In wb.Worksheets
    If sh.Name Like "Sheet *" Then
        Email_Subject = sh.Range("C129").Value
        Email_Subject = Replace(Email_Subject, "BSF- ", "")
        Email_Subject = WorksheetFunction.Proper(Email_Subject)
        
        Email_Send_From = "*INSERT EMAIL HERE*"
        Email_Send_To = sh.Range("U6").Value 
        Email_Cc = sh.Range("U7").Value 
        'Email_Body = sh.Range("C124").Value
        Email_Attach = ThisWorkbook.Path & "\" & sh.Range("R3").Value & ".pdf"
        'Email_Picture = ThisWorkbook.Path & "\" & sh.Range("R3").Value & ".bmp"
        
        On Error GoTo debugs
        
        body1 = sh.Range("C124").Value
        body2 = Replace(body1, Chr(10), "<br>") 'to add newline
        body2 = Replace(body2, "below", "attached")
        body2 = Replace(body2, "BSF-", "")
        Email_Body = body2
        
        'htmlTemp = "<!DOCTYPE html><html><body>"
        'htmlTemp = "<div id=email_body style='font-size: 12px; font-style: Arial'>"
        
        htmlTemp = "<div id=email_body>"
        htmlTemp = htmlTemp & Email_Body
        htmlTemp = htmlTemp & "<br><img src=""cid:" & sh.Range("R3").Value & ".bmp""><br>"
        htmlTemp = htmlTemp & "<br>Best Regards,"
        htmlTemp = htmlTemp & "</div>"
        
        'adding Signature
        htmlTemp = htmlTemp & "<span>--<br></span><span style='color: grey; font-family: Helvetica, sans-serif;'>YOUR NAME<br></span>"
        
        'htmlTemp = htmlTemp & "</body></html>"
        
        Set Mail_Object = CreateObject("Outlook.Application")
        Set Mail_Single = Mail_Object.CreateItem(0)
        
        'adding and embedding report image
        Set colAttach = Mail_Single.attachments
        Set oAttach = colAttach.Add(ThisWorkbook.Path & "\" & sh.Range("R3").Value & ".bmp")
        Set olkPA = oAttach.PropertyAccessor
        
        olkPA.SetProperty PR_ATTACH_CONTENT_ID, sh.Range("R3").Value & ".bmp"
        
        Mail_Single.Close olSave
        
        With Mail_Single
            .Subject = Email_Subject
            .To = Email_Send_To
            .CC = Email_Cc
            '.body = htmlTemp
            .HTMLBody = htmlTemp
            .send
        End With
        
        Set Mail_Single = Nothing
        Set colAttach = Nothing
        Set oAttach = Nothing
        Set Mail_Object = Nothing
        
debugs:
        If Err.Description <> "" Then MsgBox Err.Description
        
        
    End If
Next
    
    
End Sub

Sub ExportBitmaps()

Dim wb As Workbook
Dim sh As Worksheet

Dim PrintRange As Range
Set PrintRange = Range("A1:O120")

Dim bitmap_name As String

Set wb = ThisWorkbook

'Set sht = wb.Worksheets("Sheet 13")

'ExportRange sht.Range("A1:O120"), ThisWorkbook.Path & "\" & sht.Range("R3").Value & ".bmp"

For i = 1 To Sheets.Count
    If Sheets(i).Name Like "Sheet *" Then
        Set sh = wb.Worksheets(Sheets(i).Name)
        
        ExportRange sh.Range("A1:O120"), ThisWorkbook.Path & "\" & sh.Range("R3").Value & ".bmp"
    End If
Next

End Sub


Sub ExportRange(rng As Range, sPath As String)

    Dim cob, sc

    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    Set cob = rng.Parent.ChartObjects.Add(10, 10, 200, 200)
    'remove any series which may have been auto-added...
    Set sc = cob.Chart.SeriesCollection
    Do While sc.Count > 0
        sc(1).Delete
    Loop

    With cob
        .ShapeRange.Line.Visible = msoFalse  '<<< remove chart border
        .Height = rng.Height
        .Width = rng.Width
        .Chart.Paste
        .Chart.Export Filename:=sPath, Filtername:="PNG"
        .Delete
    End With

End Sub


Sub UnhideAll(sheet As Worksheet) 'unhide all columns so they could be hidden properly
    sheet.Rows("26:118").EntireRow.Hidden = False
End Sub

Sub HideRows()
    'this function hides all rows that are after the report date
    
    Dim wb As Workbook
    Dim sh As Worksheet
    
    Dim repD, lastD As Date
    
    Set wb = ThisWorkbook
    
    For Each sh In wb.Worksheets
        
        If sh.Name Like "Sheet *" Then
        
            UnhideAll sh
        
            repD = sh.Range("R1").Value
            For i = 26 To 118
                lastD = sh.Cells(i, 3).Value
                If lastD > repD Then
                    sh.Rows(i).EntireRow.Hidden = True
                End If
            Next
        End If
    Next
    
    'Set sh = ThisWorkbook.Sheets("Sheet 14")
    'UnhideAll sh
    
End Sub

Sub ShowLateFeeInfo()

Dim wb As Workbook
Dim sh As Worksheet

Set wb = ThisWorkbook
    
For Each sh In wb.Worksheets
    If sh.Name Like "Sheet *" Then
        
        If sh.Range("R30").Value > 33 Then 'total days loan has run
            sh.Rows(17).EntireRow.Hidden = False
            sh.Columns(12).EntireColumn.Hidden = False 'for column L with compound late fee
        Else
            sh.Rows(17).EntireRow.Hidden = True
            sh.Columns(12).EntireColumn.Hidden = True 'for column L with compound late fee
        End If
    End If
Next

End Sub

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     WorksheetExists = Not sht Is Nothing
End Function

Sub MakeSheets()

Application.Calculation = xlCalculateManual

Dim wb As Workbook
Dim sh As Worksheet

Dim shName As String
Dim prevSh As String


Set wb = ThisWorkbook

If WorksheetExists("Sheet 1", wb) = False Then 'sheet 1, our template sheet, doesn't exist
    MsgBox "MakeSheets function requires a *Sheet 1* "
Else
    For i = 2 To 20
        prevSh = "Sheet " & i - 1
        shName = "Sheet " & i
        If WorksheetExists(shName, wb) = False Then 'sheet doesn't exist, so create it
            Sheets("Sheet 1").Copy after:=Sheets(prevSh)
            ActiveSheet.Name = shName
            ActiveSheet.Range("R1").Formula = "='Data '!C" & i + 1
            ActiveSheet.Range("R2").Formula = "='Data '!H" & i + 1
            ActiveSheet.Range("R3").Formula = "='Data '!D" & i + 1
            ActiveSheet.Range("R4").Formula = "='Data '!I" & i + 1
            ActiveSheet.Range("R5").Formula = "='Data '!J" & i + 1
            ActiveSheet.Range("R6").Formula = "='Data '!L" & i + 1
            ActiveSheet.Range("R7").Formula = "='Data '!E" & i + 1
            ActiveSheet.Range("R8").Formula = "='Data '!M" & i + 1
            ActiveSheet.Range("R9").Formula = "='Data '!N" & i + 1
            ActiveSheet.Range("R10").Formula = "='Data '!O" & i + 1
            ActiveSheet.Range("R11").Formula = "='Data '!K" & i + 1
            ActiveSheet.Range("R12").Formula = "='Data '!P" & i + 1
            ActiveSheet.Range("R13").Formula = "='Data '!Q" & i + 1
            ActiveSheet.Range("R14").Formula = "='Data '!B" & i + 1
            ActiveSheet.Range("R16").Formula = "='Data '!S" & i + 1
            ActiveSheet.Range("R17").Formula = "='Data '!R" & i + 1
            ActiveSheet.Range("R18").Formula = "='Data '!T" & i + 1
            ActiveSheet.Range("R28").Formula = "='Data '!Q" & i + 1
            ActiveSheet.Range("U6").Formula = "='Data '!U" & i + 1
            ActiveSheet.Range("U7").Formula = "='Data '!V" & i + 1
        End If
    Next
    
    'wb.Sheets.Select
    'ActiveSheet.Calculate
    
End If


End Sub



