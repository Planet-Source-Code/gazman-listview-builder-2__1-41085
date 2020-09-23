Attribute VB_Name = "Module1"
Option Explicit

'************************************************************************
' Hi all, the following code is a sample of how you can work with
' ListViews. I love the 'Drag and Drop' functionality and would like to
' see someone come up with a nice and funky app' to make some real use of it.
' So make good use of the code and leave some notes on what you think, plus
' any suggestions etc.
'
' Cheers and regards,
'
' The GazMan November 2002
'
'************************************************************************

Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   lParam As Any) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' ListView functions
Public Type LVHITTESTINFO
  pt As POINTAPI
  Flags As LVHITTESTINFO_flags
  iItem As Long
#If (WIN32_IE >= &H300) Then
  iSubItem As Long
#End If
End Type

Public Enum LVHITTESTINFO_flags
  LVHT_NOWHERE = &H1   ' in LVW client area, but not over item
  LVHT_ONITEMICON = &H2
  LVHT_ONITEMLABEL = &H4
  LVHT_ONITEMSTATEICON = &H8
  LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

  'Outside the LVW's client area
  LVHT_ABOVE = &H8
  LVHT_BELOW = &H10
  LVHT_TORIGHT = &H20
  LVHT_TOLEFT = &H40
End Enum

Public Const LVM_FIRST = &H1000
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Public Const LVM_HITTEST = (LVM_FIRST + 18)

Public cnList               As ADODB.Connection
Public AccessApp            As String
Public sSQL                 As String


'************************************************************************
' The following two func's find the position of the drop...
' The GazMan November 2002
'************************************************************************

Public Function ListView_GetItemPosition(hwndLV As Long, i As Long, ppt As POINTAPI) As Boolean

  ListView_GetItemPosition = SendMessage(hwndLV, LVM_GETITEMPOSITION, ByVal i, ppt)
  
End Function

Public Function ListView_HitTest(hwndLV As Long, pinfo As LVHITTESTINFO) As Long

  ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pinfo)
  
End Function

'************************************************************************                                                                   *
' This is a generic function that creates a ListView with headers       *
' It's the "Who's yer Daddy" of ListView builders...                    *
' You just need to supply a Recordset (I'm using ADO 2.6)               *
' and a ListView to populate..                                          *
'                                                                       *
' The Gazman - November 2002                                            *
' Kia Kaha!                                                             *
'************************************************************************
Public Function BuildLVW(rsListView As Recordset, oListView As ListView)
Dim chColHead       As ColumnHeader
Dim itmNewLine      As ListItem
Dim intIndex        As Integer
Dim intCount2       As Integer
Dim intTotCount     As Integer
Dim isOkay          As Boolean

On Error GoTo Err_Handler

With rsListView
    'Clear out the ListView ready for new values....
    oListView.ColumnHeaders.Clear
    oListView.ListItems.Clear
    'Set up the column headers and then add the items to the ListView control....
    
    'First, Add the Column Headers
    For intIndex = 0 To .Fields.Count - 1
        Set chColHead = oListView.ColumnHeaders.Add(, , rsListView(intIndex).Name)
        oListView.ColumnHeaders.Item(intIndex + 1).Width = 2000 'Adjust size if needed...
    Next intIndex
    
    ' Now, loop through the recordset and add Items to the lvw.....
    If .EOF = False Then
        intTotCount = rsListView.RecordCount
        For intIndex = 1 To intTotCount
            If IsNull(rsListView(0).Value) = False Then
                If IsNumeric(rsListView(0).Value) Then
                    Set itmNewLine = oListView.ListItems.Add(, , Str(rsListView(0).Value))
                Else
                     Set itmNewLine = oListView.ListItems.Add(, , rsListView(0).Value)
                End If
                
                For intCount2 = 1 To .Fields.Count - 1 'Add the subitems..
                    If IsNull(rsListView(intCount2).Value) = False Then
                        itmNewLine.SubItems(intCount2) = rsListView(intCount2).Value
                    Else
                        itmNewLine.SubItems(intCount2) = ""
                    End If
                Next intCount2
                .MoveNext
            Else
                'If the value IsNull then must add a blank...otherwise muchos problemos
                '(try it without this line...must have data with Nulls on board)
                Set itmNewLine = oListView.ListItems.Add(, , "")
            End If
        Next intIndex
        oListView.GridLines = True
    Else
        oListView.ListItems.Clear
        oListView.GridLines = False
    End If
End With

Exit Function
Err_Handler:

    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

'************************************************************************
' Updates back to the database from the items in the listview....
'
' The GazMan November 2002
'
'************************************************************************
Function Update(lvwList As ListView, sTable As String)

Dim rsUpdate        As ADODB.Recordset
Dim entry           As ListItem
Dim strDocCode      As String
Dim sSQL            As String
Dim i               As Integer

On Error GoTo Err_Handler

'Clear out the table ready for all the new values....
sSQL = "DELETE * From " & sTable & ""
cnList.Execute sSQL

'Add the values from the ListView....
Set rsUpdate = New ADODB.Recordset
sSQL = "SELECT * From " & sTable & ""
rsUpdate.Open sSQL, cnList, adOpenStatic, adLockOptimistic

With rsUpdate
    i = 1
    For Each entry In lvwList.ListItems
        .AddNew

        .Fields("ID") = lvwList.ListItems.Item(i)
        .Fields("CampID") = lvwList.ListItems.Item(i).ListSubItems(1)
        .Fields("TableName") = lvwList.ListItems.Item(i).ListSubItems(2)
        .Fields("Alias") = lvwList.ListItems.Item(i).ListSubItems(3)
        .Fields("MergeField") = lvwList.ListItems.Item(i).ListSubItems(4)
             
        .Update
        i = i + 1
    Next
End With

rsUpdate.Close
Set rsUpdate = Nothing

Exit Function
Err_Handler:

    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

'************************************************************************
' Populates the textboxes from the selected item in the listview
' Uses the text1 array of textboxes and the selected item from the
' relevant listview...
'
' The GazMan November 2002
'
'************************************************************************
Public Function SetTextBoxes(lvwPacks As ListView, frmForm As Form)

Dim sTxbName        As String
Dim intTotCount     As Integer
Dim intIndex        As Integer
Dim intItem         As Integer

On Error GoTo Err_Handler

With frmForm
    intTotCount = lvwPacks.ColumnHeaders.Count
    intItem = 0
    For intIndex = 1 To intTotCount
        If intIndex = 1 Then
            If lvwPacks.SelectedItem Is Nothing Then .txbID.Text = "" Else
                .Text1(intItem).Text = Trim(lvwPacks.SelectedItem)
        Else
            If lvwPacks.SelectedItem Is Nothing Then .txbID.Text = "" Else
                .Text1(intItem).Text = Trim(lvwPacks.SelectedItem.SubItems(intItem))
        End If
        intItem = intItem + 1
    Next intIndex
End With

Exit Function

Err_Handler:

    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation
    
End Function


'************************************************************************
' Example of creating a SQL string with variable from the ListView
'
' The GazMan November 2002
'
'************************************************************************
Function RecordDetails(lvwList As ListView, sTable As String)

Dim rsUpdate        As ADODB.Recordset
Dim sSQL            As String
Dim sListItem       As String

On Error GoTo Err_Handler

' Could also use....
' lvwList.SelectedItem.SubItems (2)
sListItem = Trim(lvwList.SelectedItem)

'Add the values from the ListView....
Set rsUpdate = New ADODB.Recordset
sSQL = "SELECT * From " & sTable & " Where ID = " & sListItem & ""
rsUpdate.Open sSQL, cnList, adOpenStatic, adLockOptimistic

MsgBox "Record returned from lvwDataCopy: " & vbCrLf & vbCrLf & "" _
        & rsUpdate.Fields("ID") & vbCrLf _
        & rsUpdate.Fields("CampID") & vbCrLf _
        & rsUpdate.Fields("TableName") & vbCrLf _
        & rsUpdate.Fields("Alias") & vbCrLf _
        & rsUpdate.Fields("MergeField") & vbCrLf, vbInformation, "ListView"
       
rsUpdate.Close
Set rsUpdate = Nothing
    
Exit Function
Err_Handler:

    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

