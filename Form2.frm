VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form2"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12105
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8400
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDetailsDataCopy 
      Caption         =   "Details"
      Height          =   375
      Left            =   10680
      TabIndex        =   19
      Top             =   2820
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Selected Record"
      Height          =   1935
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   11295
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   13
         Text            =   "txbAlias"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   4
         Left            =   3960
         TabIndex        =   12
         Text            =   "txbMergeField"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   11
         Text            =   "txbTablename"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   10
         Text            =   "txbID"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Text            =   "txbCampID"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "MergeField :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alias :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "TableName :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "CampID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10680
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwDataCopy 
      Height          =   4965
      Left            =   6120
      TabIndex        =   0
      Top             =   3240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8758
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483646
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   4965
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8758
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483646
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "lvwDataCopy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "lvwData"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ListView Builder 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      Height          =   8175
      Left            =   120
      Top             =   120
      Width           =   11760
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DragLV                      As ListItem 'The item being dragged

Private Sub cmdClose_Click()

    Unload Me
    
End Sub


Private Sub cmdDetailsDataCopy_Click()

    RecordDetails lvwDataCopy, "tbDataCopy"

End Sub

Private Sub cmdUpdate_Click()

    Update lvwData, "tbData"
    Update lvwDataCopy, "tbDataCopy"

End Sub

Private Sub Command1_Click()

    SelectRecords lvwData, "tbData"
    SelectRecords lvwDataCopy, "tbDataCopy"
    
End Sub

Private Sub Form_Load()
Dim i       As Integer

On Error GoTo Err_Handler

' Set up the Connection to the Access database....
AccessApp = Trim(App.Path & "\Data.mdb")
Set cnList = New ADODB.Connection
cnList.Open "PROVIDER=MSDASQL;" & _
        "DRIVER={Microsoft Access Driver (*.mdb)};" & _
              "DBQ= " & AccessApp & ";" & _
              "UID=sa;PWD=;"

'Clear the text boxes...
i = 0
Do Until i = Text1.Count
    Text1(i).Text = ""
    i = i + 1
Loop

' Populate the ListViews....
Command1_Click

Exit Sub
Err_Handler:
        
    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation
    
End Sub

'************************************************************************
' Populate the ListViews....
'
' The GazMan November 2002
'
'************************************************************************
Sub SelectRecords(lvwList As ListView, sTable As String)
Dim rsMergeSelect           As ADODB.Recordset

On Error GoTo Err_Handler

Set rsMergeSelect = New ADODB.Recordset
sSQL = "SELECT * From " & sTable & "  Order by ID"
rsMergeSelect.Open sSQL, cnList, adOpenStatic, adLockReadOnly

If rsMergeSelect.EOF = False Then BuildLVW rsMergeSelect, lvwList

rsMergeSelect.Close
Set rsMergeSelect = Nothing

Exit Sub
Err_Handler:
        
    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Sub

'************************************************************************
' Grabs the selected item from the ListView
'
' The GazMan November 2002
'
'************************************************************************
Private Sub lvwData_ItemClick(ByVal Item As MSComctlLib.ListItem)

    SetTextBoxes lvwData, Form2

End Sub

'************************************************************************
' Enables the Drag and Drop functionality between the two listviews..
' This one is from lvwData to lvwDataCopy.... (oooww, I like this...)
'
' The GazMan November 2002
'
'************************************************************************
Private Sub lvwData_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, _
    Button As Integer, Shift As Integer, x As Single, y As Single)

Dim liNew       As ListItem
Dim pinfo       As LVHITTESTINFO
Dim pt          As POINTAPI
Dim pti         As POINTAPI
Dim hitItem     As ListItem
Dim i           As Integer
Dim bNew        As Boolean
   
On Error GoTo Err_Handler

Set hitItem = lvwData.HitTest(x, y)

If Not hitItem Is Nothing Then

    Set liNew = lvwData.ListItems.Add(hitItem.Index, , DragLV.Text)
    i = 1
    Do Until i = DragLV.ListSubItems.Count + 1
        liNew.SubItems(i) = DragLV.SubItems(i)
        i = i + 1
    Loop
    liNew.Selected = True
    lvwDataCopy.ListItems.Remove DragLV.Index
        
Else
    
    GetCursorPos pt
    
    If lvwData.ListItems.Count < 2 Then
        bNew = True
        Set liNew = lvwData.ListItems.Add(, , DragLV.Text)
        Call ListView_GetItemPosition(lvwDataCopy.hwnd, _
            lvwData.ListItems.Item(lvwData.ListItems.Count).Index, pti)
    Else
    
        Call ListView_GetItemPosition(lvwData.hwnd, _
            lvwData.ListItems.Item(lvwData.ListItems.Count - 1).Index, pti)
    End If
    
    If pt.y > Me.Top / Screen.TwipsPerPixelY + pti.y Then
        If bNew = False Then Set liNew = lvwData.ListItems.Add(, , DragLV.Text)
            i = 1
            Do Until i = DragLV.ListSubItems.Count + 1
                liNew.SubItems(i) = DragLV.SubItems(i)
                i = i + 1
            Loop
            liNew.Selected = True
        lvwDataCopy.ListItems.Remove DragLV.Index
    End If
    
End If
    
Exit Sub
Err_Handler:
        
    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Sub

'************************************************************************
' Start of the Drag, the item is ID'd and set to DragLV
'
' The GazMan November 2002
'
'************************************************************************
Private Sub lvwData_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    
    Set DragLV = lvwData.SelectedItem

End Sub

'************************************************************************
' Grabs the selected item from the ListView
'
' The GazMan November 2002
'
'************************************************************************
Private Sub lvwDataCopy_ItemClick(ByVal Item As MSComctlLib.ListItem)

    SetTextBoxes lvwDataCopy, Form2

End Sub

'************************************************************************
' Enables the Drag and Drop functionality between the two listviews..
' This one is from lvwDataCopy to lvwData ....
'
' The GazMan November 2002
'
'************************************************************************
Private Sub lvwDataCopy_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, _
                    Button As Integer, Shift As Integer, x As Single, y As Single)

Dim liNew       As ListItem
Dim pinfo       As LVHITTESTINFO
Dim pt          As POINTAPI
Dim pti         As POINTAPI
Dim hitItem     As ListItem
Dim i           As Integer
Dim bNew        As Boolean
   
On Error GoTo Err_Handler

Set hitItem = lvwDataCopy.HitTest(x, y)

If Not hitItem Is Nothing Then

    Set liNew = lvwDataCopy.ListItems.Add(hitItem.Index, , DragLV.Text)
    i = 1
    Do Until i = DragLV.ListSubItems.Count + 1
        liNew.SubItems(i) = DragLV.SubItems(i)
        i = i + 1
    Loop
    liNew.Selected = True
    lvwData.ListItems.Remove DragLV.Index
        
Else
    
    GetCursorPos pt
    
    If lvwDataCopy.ListItems.Count < 2 Then
        bNew = True
        Set liNew = lvwDataCopy.ListItems.Add(, , DragLV.Text)
        Call ListView_GetItemPosition(lvwDataCopy.hwnd, _
            lvwDataCopy.ListItems.Item(lvwDataCopy.ListItems.Count).Index, pti)
    Else
        Call ListView_GetItemPosition(lvwDataCopy.hwnd, _
            lvwDataCopy.ListItems.Item(lvwDataCopy.ListItems.Count - 1).Index, pti)
    End If
    
    If pt.y > Me.Top / Screen.TwipsPerPixelY + pti.y Then
        If bNew = False Then Set liNew = lvwDataCopy.ListItems.Add(, , DragLV.Text)
            i = 1
            Do Until i = DragLV.ListSubItems.Count + 1
                liNew.SubItems(i) = DragLV.SubItems(i)
                i = i + 1
            Loop
            liNew.Selected = True
        lvwData.ListItems.Remove DragLV.Index
    End If
    
End If
    
Exit Sub
Err_Handler:
        
    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Sub

'************************************************************************
' Start of the Drag, the item is ID'd and set to DragLV
'
' The GazMan November 2002
'
'************************************************************************

Private Sub lvwDataCopy_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

    Set DragLV = lvwDataCopy.SelectedItem

End Sub
