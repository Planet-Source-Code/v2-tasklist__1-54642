VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task List"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnVote 
      Caption         =   "Vote For This Code"
      Height          =   435
      Left            =   135
      TabIndex        =   2
      Top             =   5640
      Width           =   2235
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      Height          =   435
      Left            =   2460
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ListBox lstTasks 
      Height          =   5520
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4875
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub WindowActivated(hwnd As Long)
Dim i As Integer

  For i = 0 To lstTasks.ListCount - 1
    If lstTasks.ItemData(i) = hwnd Then
        lstTasks.ListIndex = i
    End If
  Next

End Sub

Public Sub WindowRedraw(hwnd As Long)
Dim strCaption As String
Dim i As Integer
  
  strCaption = Space(1024)
  
  GetWindowText hwnd, strCaption, 1024
                                      
                                      
  For i = 0 To lstTasks.ListCount - 1
    If lstTasks.ItemData(i) = hwnd Then
      lstTasks.List(i) = strCaption
      lstTasks.ListIndex = i
      Exit For
    End If
  Next
End Sub

Public Sub WindowCreated(hwnd As Long)
  Dim i As Integer
  Dim lExStyle    As Long
  Dim bNoOwner    As Boolean
  Dim lreturn     As Long
  Dim sWindowText As String
  
  If Not hwnd = Me.hwnd Then
    If IsWindowVisible(hwnd) Then
        If GetParent(hwnd) = 0 Then
            bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                
                sWindowText = Space$(1024)
                lreturn = GetWindowText(hwnd, sWindowText, Len(sWindowText))
                If lreturn Then
                   sWindowText = Left$(sWindowText, lreturn)
                     lstTasks.AddItem sWindowText
                     lstTasks.ItemData(lstTasks.NewIndex) = hwnd
                     Caption = "Task List - [ " & lstTasks.ListCount & " Task(s) Running. ]"
                End If
            End If
        End If
    End If
  End If
  
  
End Sub

Public Sub WindowDestroyed(hwnd As Long)
  Dim i As Integer
  
  For i = 0 To lstTasks.ListCount - 1 'Loop around looking for the hwnd and
                                   '  remove it from the list
    If lstTasks.ItemData(i) = hwnd Then
      lstTasks.RemoveItem i
      Exit For
    End If
  Next
Caption = "Task List - [ " & lstTasks.ListCount & " Task(s) Running. ]"
End Sub

Private Sub btnVote_Click()
    ShellExecute 0, "OPEN", "http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=54642&lngWId=1", "", "", 1
End Sub

Private Sub btnQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  SetMM
  SetHook Install
  MsgBox "Try Running New Tasks And Then End Them To See The Results.", vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetHook Release
End Sub
