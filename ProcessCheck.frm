VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Sniffer"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Refresh ListBoxes"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Query (q)"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Query Checking"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   7560
      Width           =   5175
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove"
         Height          =   300
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check all"
         Height          =   300
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
      Begin VB.ListBox List3 
         Height          =   645
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   5175
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "List System Processes"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Opened windows are Max, Focused"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   4935
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remove from list (r)"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2535
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00C0C0FF&
      Height          =   4740
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remove from list (r)"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2535
   End
   Begin VB.ListBox processes 
      Height          =   255
      ItemData        =   "ProcessCheck.frx":0000
      Left            =   5520
      List            =   "ProcessCheck.frx":0002
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFC0&
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Not indentified processes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Processes in Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROCESS SNIFFER
'Made by: Mephisto

'forgive me if i dont go to very detail on explaining, but i posted this for experts. So i
'assume you know basic functions, listboxes, and basic way to interact with API's

'getting processes API's
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'normal variables used
Dim arr() As String
Dim StrFileIn As String
Dim i As Integer
Dim j As Integer
Dim num As Integer

Dim process() As String
Dim pNum As Integer
Dim match As Boolean

Private Sub Command1_Click()
'remove red listbox entry
List2.RemoveItem List2.ListIndex
End Sub

Private Sub Command2_Click()
'remove green listbox entry
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
'for Query checking. Basically here what we are doing is we are setting up a loop that does
'the same thing as Check button.

For i = 0 To List3.ListCount - 1

Dim sel As String

'we get the process name
sel = List3.List(i)
'replace .exe
sel = Replace(sel, ".exe", "")
'or if there is big .EXE remove that
sel = Replace(sel, ".EXE", "")
'remove all comments
sel = Replace(sel, " - ATTENTION", "")
sel = Replace(sel, " - SYSTEM", "")

'we test whether the max focus check is checked
If Check1.Value = 1 Then
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE www.liutilities.com/products/wintaskspro/processlibrary/" & sel, vbMaximizedFocus
Else
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE www.liutilities.com/products/wintaskspro/processlibrary/" & sel
End If

Next i
End Sub

Private Sub Command4_Click()
'To check for the process
Dim sel As String

'filter the crap so we are left with pure process name
sel = List1.List(List1.ListIndex)
sel = Replace(sel, ".exe", "")
sel = Replace(sel, ".EXE", "")
sel = Replace(sel, " - ATTENTION", "")
sel = Replace(sel, " - SYSTEM", "")

'all entries on the site are very nicely done. It is just the process name. This allows me to
'make easy code by simply taking the base and then adding on the process name and sending the
'internet explorer in there.
If Check1.Value = 1 Then
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE www.liutilities.com/products/wintaskspro/processlibrary/" & sel, vbMaximizedFocus
Else
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE www.liutilities.com/products/wintaskspro/processlibrary/" & sel
End If

End Sub

Private Sub Command5_Click()
Dim sel As String
'when user clicks on red listbox search, we launch google and put the process name in

sel = List2.List(List2.ListIndex)
sel = LCase(sel)

If Check1.Value = 1 Then
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://www.google.ca/search?q=""" & sel & """&ie=UTF-8&oe=UTF-8&hl=en&meta=", vbMaximizedFocus
Else
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://www.google.ca/search?q=""" & sel & """&ie=UTF-8&oe=UTF-8&hl=en&meta="
End If
End Sub

Private Sub Command6_Click()
List3.RemoveItem (List3.ListIndex)
End Sub

Private Sub Command7_Click()
List3.AddItem List1.List(List1.ListIndex)
End Sub

Private Sub Command9_Click()
'we go from beggining
Call Form_Load
End Sub

Private Sub Form_Load()
'initialize all listboxes and stuff
List1.Clear
List2.Clear
processes.Clear

'processes listbox is invisible on the form. I am aware that this listbox is pretty useless and can
'be left out, but in early development of this program i used it, so i wont remove it now, i am
'afraid i will create bugs. Anyway, it doesnt hurt noone :D

GetProcesses
SortProcesses
End Sub

Sub GetProcesses()
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    'Takes a snapshot of the processes and the heaps, modules, and threads used by the processes
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    'set the length of our ProcessEntry-type
    uProcess.dwSize = Len(uProcess)
    'Retrieve information about the first process encountered in our system snapshot
    r = Process32First(hSnapShot, uProcess)
    'set graphics mode to persistent
    Me.AutoRedraw = True
    Do While r
        'Retrieve information about the next process recorded in our system snapshot
        r = Process32Next(hSnapShot, uProcess)
        processes.AddItem Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
    Loop
    'close our snapshot handle
    CloseHandle hSnapShot
    
    'we filter out some more scam
    processes.RemoveItem (0)
    processes.RemoveItem (processes.ListCount - 1)
End Sub

Private Sub List1_DblClick()
'equal to check
Call Command4_Click
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
'for shortcuts, we use this code.
If KeyAscii = Asc("q") Then
    Call Command7_Click
ElseIf KeyAscii = Asc("r") Then
    Call Command2_Click
End If
End Sub

Private Sub List2_DblClick()
Call Command5_Click
End Sub

Sub SortProcesses()
Dim temp() As String
Dim temp2() As String

List1.Clear

'loop through the list of all processes available and store them in array arr()
For i = 0 To processes.ListCount - 1
ReDim Preserve arr(num)
arr(num) = processes.List(num)
num = num + 1
Next i

'we now read in all processes names that are in database
Open App.Path & "/processes.txt" For Binary As #1
StrFileIn = Input$(LOF(1), 1)
Close #1

'now the tricky part., we have to extract the processes
'first we seprate the text by lines and store them all in temp() array
temp = Split(StrFileIn, vbCrLf)

For i = 0 To UBound(temp)
'now we separate each temp() by 3 spaces and read them into temp2() array
temp2 = Split(temp(i), "   ")

'the first word there
ReDim Preserve process(pNum)
process(pNum) = temp2(0)
pNum = pNum + 1

'second word
ReDim Preserve process(pNum)
process(pNum) = temp2(1)
pNum = pNum + 1

'third word
ReDim Preserve process(pNum)
process(pNum) = temp2(2)
pNum = pNum + 1

'to make sure it wont cycle the PC to death
DoEvents
Next i

'to make sure all processes are without unnecessary spaces and all
'there are no processes that have spaces in them
For i = 0 To UBound(process)
process(i) = Replace(process(i), " ", "")
Next i

For i = 0 To UBound(arr)
    match = False
    For j = 0 To UBound(process)
        If LCase(arr(i)) = LCase(process(j)) Then
        match = True
        'here we search for the processes. this is the core of the program. we are looking
        'which processes are in database and which arent. We put them in LCase to make sure
        'it is not case sensitive
        Exit For
        End If
    Next j
    
    If match = True Then
        If j < 125 Then
        'first 126 processes are Viruses/Spyware/Adware/ etc
        List1.AddItem arr(i) & " - ATTENTION"
        ElseIf j < 207 Then
        'next processes (126-208) are System processes
            If Check2.Value = 1 Then List1.AddItem arr(i) & " - SYSTEM"
        Else
        'other are just simple application processes
        List1.AddItem arr(i)
        End If
    
    Else
    'if the match is false, so that the process is not found, add it to red listbox
    List2.AddItem arr(i)
    End If
    
    'give windows some time to do other stuff
    DoEvents
Next i

End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("r") Then
    Call Command1_Click
End If
End Sub
