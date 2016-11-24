VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "Rename"
      Height          =   465
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   5730
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save File"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   4590
      Left            =   3960
      TabIndex        =   2
      Top             =   870
      Width           =   3765
   End
   Begin VB.DirListBox Dir1 
      Height          =   4500
      Left            =   300
      TabIndex        =   1
      Top             =   870
      Width           =   3585
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   330
      TabIndex        =   0
      Top             =   360
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
   Unload Me
End Sub



Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub Command1_Click()
   Dim i As Long
   Dim sFileName As String
   
   For i = 0 To File1.ListCount - 1
      sFileName = Trim(File1.List(i))
      If UCase(Right(sFileName, 4)) = ".MP3" Then
         'Debug.Print Left(sFileName, Len(sFileName) - 4)
         WriteToTxtFile App.Path & "\music.txt", Left(sFileName, Len(sFileName) - 4)
      End If
   Next
   
   MsgBox "Success!", vbInformation, "OK"
End Sub

'Rename
Private Sub Command3_Click()
   Dim i As Long
   Dim sFileName As String
   Dim sNewName As String
   
   Dim sPathName As String
   Dim sNewPathName As String
   
   For i = 0 To File1.ListCount - 1
      sFileName = Trim(File1.List(i))
      'If UCase(Right(sFileName, 4)) = ".MP3" Then
         
'''         '显示第4位不是点和空格的文件
'''         If Mid(sFileName, 4, 1) <> "." Then
'''            Debug.Print sFileName
'''         End If
         If i = 100 Or i = 200 Then
            Debug.Print " "
         End If
         
         sNewName = Trim(Mid(sFileName, 5))   '从原文件名的第5位开始取

         If Left(sNewName, 1) = "." Then sNewName = Mid(sNewName, 2)   '如果第1位是“.”  则去掉
         sNewName = Trim(sNewName)

         sNewName = Right("000" & (i + 1), 3) & "." & sNewName   '前面加上“00X.”   001.歌名.mp3
         'Debug.Print sNewName

         sPathName = Dir1.Path & "" & sFileName
         sNewPathName = Dir1.Path & "" & sNewName

         Debug.Print sNewPathName
         Name sPathName As sNewPathName
      'End If
   Next
   
   MsgBox "Success!", vbInformation, "OK"
   
End Sub

Private Sub WriteToTxtFile(ByVal sFileName As String, ByVal strTmp As String)
On Error GoTo Error
   Dim fnum As Long, FL As Long
      
'   If Dir(sFileName) <> "" Then
'      FL = FileLen(sFileName)
'      If FL / 1024 / 1024 / 1024 > 1 Then   '大于10M，备份文件，删除原文件
'         FileCopy sFileName, Left(sFileName, (Len(sFileName) - 4)) & Format(Now, "yyyyMMddhhmmss") & ".txt"
'         Kill sFileName
'      End If
'   End If
   
   fnum = FreeFile
   Open sFileName For Append As #fnum
''   FL = LOF(fnum)
''   If FL / 1024 / 1024 / 1024 > 1 Then   '大于10M退出，并备份文件，删除原文件
''      Close #fnum
''      FileCopy sFileName, Left(sFileName, (Len(sFileName) - 4)) & Format(Now, "yyyyMMddhhmmss") & ".txt"
''      Kill sFileName
''      Exit Sub
''   End If
   
   Print #fnum, strTmp
   Close #fnum
   Exit Sub
Error:
   On Error Resume Next
   Close #fnum
End Sub
