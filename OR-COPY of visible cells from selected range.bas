Attribute VB_Name = "Module1"
'Handle 64-bit and 32-bit Office
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, _
    ByVal dwBytes As LongPtr) As LongPtr
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
  Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As LongPtr
  Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, _
    ByVal hMem As LongPtr) As LongPtr

Const GHND = &H42
Const CF_TEXT = 1
Const MAXSIZE = 4096

Function ClipBoard_SetData(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

Dim hGlobalMemory As Long, lpGlobalMemory As Long
Dim hClipMemory As Long, X As Long

'Allocate moveable global memory
hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

'Lock the block to get a far pointer to this memory.
lpGlobalMemory = GlobalLock(hGlobalMemory)

'Copy the string to this global memory.
lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

'Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
MsgBox "Could not unlock memory location. Copy aborted."
    GoTo OutOfHere2
  End If

'Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
MsgBox "Could not open the Clipboard. Copy aborted."
    Exit Function
  End If

'Clear the Clipboard.
X = EmptyClipboard()

'Copy the data to the Clipboard.
hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
  If CloseClipboard() = 0 Then
MsgBox "Could not close Clipboard."
  End If

End Function

Sub CopyTextToClipboard()
'PURPOSE: Copy a given text to the clipboard (using Windows API)
'SOURCE: www.TheSpreadsheetGuru.com
'NOTES: Must have above API declaration and ClipBoard_SetData function in your code

'Place text into the Clipboard
ClipBoard_SetData disjunctionConcatVisibleTextOfRange()

End Sub

' @author tberna
' concatenate text from selected range, visible cells only, with OR between items. "a OR b OR c ..."
Function disjunctionConcatVisibleTextOfRange() As String

    Dim cel As Range
    Dim selectedRange As Range
    Dim output As String

    Set selectedRange = Application.Selection

    For Each cel In selectedRange.Cells
        If Not ActiveSheet.Rows(cel.Row).Hidden Then
            output = output + cel.Text + " OR "
        End If
    Next cel
    
    output = Left(output, Len(output) - 4)
    disjunctionConcatVisibleTextOfRange = output

End Function

