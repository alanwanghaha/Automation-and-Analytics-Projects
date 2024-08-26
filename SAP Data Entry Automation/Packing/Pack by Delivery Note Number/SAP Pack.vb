Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const BM_CLICK = &HF5

'---
Sub SAP_Pack()

'Run Info_Check for error proof
If Info_Check = False Then
    Exit Sub
End If

Application.ScreenUpdating = False


Worksheets("Template1").Activate

Dim RowNumber As Integer
'Count number of rows by DN No. from K2
Cells(2, 11).Activate
RowNumber = Cells(Rows.Count, (ActiveCell.Column)).End(xlUp).Row
Debug.Print ("RowNumber: " & RowNumber)

Dim DN_No As String
Dim NxtDN_No As String
Dim LastRow_DN As Integer
Dim Qty As Integer
Dim index As Integer
Dim ws As Worksheet
Set ws = Worksheets("Template1")
Dim DN_List as String
DN_List = ""

'Activate SAP module - VL02N
If Not IsObject(guiApplication) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set guiApplication = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(Connection) Then
    Set Connection = guiApplication.Children(0)
End If

If Not IsObject(session) Then
    Set session = Connection.Children(0)
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject guiApplication, "on"
End If
session.findById("wnd[0]").maximize

'Activate Excel Sheet Template
Dim objExcel
Dim objSheet, intRow, i
Set objExcel = GetObject(, "Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet



'Enter T code in search bar
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl02n"
session.findById("wnd[0]").sendVKey 0  'press enter


'Loop from the first row of data (starting from the third row in Excel)
'For i = 3 To RowNumber
index = 3
Do While index <= RowNumber

    LastRow_DN = index

    Do
        DN_No = Trim(CStr(objSheet.Cells(LastRow_DN, 11).Value))  'Take DN No. from K3 downwards
        Debug.Print ("DN_No: " & DN_No)

        If index = RowNumber Then
            NxtDN_No = ""
        Else
            NxtDN_No = Trim(CStr(objSheet.Cells(LastRow_DN + 1, 11).Value)) 'Take next DN#
        End If
        Debug.Print ("NxtDN_No: " & NxtDN_No)
    
    Loop Until LastRow_DN = RowNumber Or DN_No <> NxtDN_No
    Debug.Print ("LastRow_DN: " & LastRow_DN)

    Qty = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(index, 10), ws.Cells(LastRow_DN, 10)))
    Debug.Print ("Qty: " & Qty)

    'Key in DN#
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = DN_No
    session.findById("wnd[0]").sendVKey 0  'press enter
    session.findById("wnd[0]/tbar[1]/btn[18]").press  'click packing

    'Enter packing - HU
    'Loop from qty1 to the last
    If Qty < 6 Then  'if qty is 1-5 then key in row by row for each pcs, else key in one line only and and key in the Qty
        For j = 3 To Qty + 2
            On Error Resume Next
            session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").Text = j - 2
            session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").Text = "BOXM"   'packing material row 1
            session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").verticalScrollbar.Position = j - 4
            session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001").verticalScrollbar.Position = j - 2
            Debug.Print("Qty <=5")
        Next j
    Else
            session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-EXIDV[0,0]").Text = Qty
            session.findById("wnd[0]/usr/tabsTS_HU_VERP/tabpUE6POS/ssubTAB:SAPLV51G:6010/tblSAPLV51GTC_HU_001/ctxtV51VE-VHILM[2,0]").Text = "BOXM"   'packing material row 1
            Debug.print("Qty > 5")
    End If


    'To print out
    'session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/btn[11]").press  'save
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = 9 'select Outbound Delivery field
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = DN_No      'Input DN# in the field 
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = 9 'select Outbound Delivery field
    session.findById("wnd[0]").sendVKey 0 'press enter to search
    session.findById("wnd[0]/mbar/menu[3]/menu[1]/menu[0]").Select  'select "Extra > Delivery Output > Header" from the top panel



    'TO LOCATE THE FIRST AVAILABEL ROW TO SET PRINTING METHOD
    'set k as the first available row to key in 'ZUCI'
    Dim k As Integer
    k = 0
    Do While True
        Dim cellText As String
        cellText = session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," & k & "]").Text
        If cellText = "" Then Exit Do
        k = k + 1
        
    Loop
    Debug.Print ("first available row: " & k)

    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," & k & "]").SetFocus  'To select the cell
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," & k & "]").Text = "ZUCI"
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," & k & "]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/usr/chkNAST-DIMME").Selected = True
    session.findById("wnd[0]/usr/chkNAST-DELET").Selected = True
    session.findById("wnd[0]/usr/ctxtNAST-LDEST").Text = "LOCAL"
    session.findById("wnd[0]/tbar[0]/btn[3]").press  'back (F3)
    session.findById("wnd[0]/tbar[0]/btn[11]").press   'save

    WaitForPrintDialog 'Call sub procudure WaitForPrintDialog to process the printing prompt
    index = LastRow_DN + 1 'To update index

    'Add processed DN_No in a list
    If DN_List = "" Then
        DN_List = DN_No  ' Add in the first DN# to the list
    Else
        DN_List = DN_List & vbCrLf & DN_No  ' Add subsequent DN# on new lines
    End If


Loop
    session.findById("wnd[0]/tbar[0]/btn[3]").press  'back (F3) to go back to the home page
    MsgBox "Printed DN Numbers:" & vbCrLf & DN_List, vbInformation, "Printed DNs"

End Sub
'---

Function Info_Check() As Boolean  
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Template1")

    Dim response As Integer
    Dim PrinterName as String
    Dim pos as Integer

    PrinterName = Application.ActivePrinter
    pos = InStr (PrinterName, "on")

    'To eliminate the printer port name if port name is found after 'on'
    If pos > 0 then
        PrinterName = Left(PrinterName, pos-1)
    End If

    response = MsgBox("Do you want to print with the default printer?" & vbNewLine & PrinterName, _
                      vbInformation + vbYesNo, "Confirm Printer")  'A yes/no msgbox to ask the user if they would like to proceed printing the document with the default printer
    If response = vbNo Then
      Info_Check = False
        MsgBox "Printing canceled."
        Exit Function ' Exit the subroutine early if the user selected cancel
    End If

    Application.ScreenUpdating = False

    ' Get final row index by DN in column K downwards
    Dim RowNumber As Integer
    RowNumber = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

    ' warn the user if no DN# is found
    If RowNumber < 3 Then 
        MsgBox "No DN# input in Template1 column K."
        Exit Function
    End If


    ' Check DN# format
    Dim i As Integer
    For i = 3 To RowNumber
        ' Check DN# are all in 9 digits.
        If Len(ws.Cells(i, 11).Value) <> 9 Then
            MsgBox "DN input in Cell K" & i & " is not in 9 digits. Please check again."
            Exit Function
        End If

        ' Check no Qty is missed out
        If ws.Cells(i, 10).Value = 0 Then
            MsgBox "Please input Qty in Cell J" & i
            Exit Function
        End If
    Next i

    Application.ScreenUpdating = True
    Info_Check = True
End Function

Sub PrintOK()
    Dim hWndPrintDialog As Long
    Dim hWndOKButton As Long

    ' Find the Printing Window
    hWndPrintDialog = FindWindow("#32770", "Print")

    If hWndPrintDialog <> 0 Then
        ' Find the OK button in the dialog
        hWndOKButton = FindWindowEx(hWndPrintDialog, 0, "Button", "OK")

        If hWndOKButton <> 0 Then
            SendMessage hWndOKButton, BM_CLICK, 0, 0  'To click the OK button
        Else
            MsgBox "OK button not found." 'If the OK button is not found, pop up a msgbox to warn the user
        End If
    Else
        MsgBox "Print dialog not found."  'If the printing prompt is not found then warn the user
    End If
End Sub

Sub WaitForPrintDialog()
    Dim hWndPrintDialog As Long

    ' Loop until the print dialog is found
    Do
        hWndPrintDialog = FindWindow("#32770", "Print")  
        If hWndPrintDialog <> 0 Then Exit Do ' Exit loop if the window is found
        
        ' Pause for 500 milliseconds
        Sleep 500  ' Maintenance Caution!! Make sure Sleep is declared in the very beginning of this script
        
        ' Allow Excel to process other events
        DoEvents
    Loop

    ' Call sub procedure PrintOK to click OK if the dialog is found
    If hWndPrintDialog <> 0 Then PrintOK
End Sub


