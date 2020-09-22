VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dir Recursion - By Deth"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Dir File Information"
      Height          =   330
      Left            =   1440
      TabIndex        =   14
      Top             =   495
      Width           =   1770
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Folder Size"
      Height          =   330
      Left            =   90
      TabIndex        =   13
      Top             =   495
      Width           =   1275
   End
   Begin VB.ListBox List2 
      Height          =   3765
      Left            =   3690
      TabIndex        =   8
      Top             =   1575
      Width           =   3525
   End
   Begin VB.Frame Frame1 
      Caption         =   "Return All"
      Height          =   735
      Left            =   3375
      TabIndex        =   5
      Top             =   540
      Width           =   3750
      Begin VB.CheckBox Check2 
         Caption         =   "Return As Folder\Files"
         Height          =   240
         Left            =   1530
         TabIndex        =   7
         Top             =   315
         Width           =   1950
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Get Dir"
         Height          =   330
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Recursion"
      Height          =   240
      Left            =   3420
      TabIndex        =   4
      Top             =   180
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Folders"
      Height          =   330
      Left            =   6030
      TabIndex        =   3
      Top             =   135
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Text            =   "c:\windows"
      Top             =   135
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Files"
      Height          =   330
      Left            =   4905
      TabIndex        =   1
      Top             =   135
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   90
      TabIndex        =   0
      Top             =   1575
      Width           =   3525
   End
   Begin VB.Label lblStatus 
      Caption         =   "Easy Folder And File Listing Code..."
      Height          =   240
      Left            =   90
      TabIndex        =   12
      Top             =   5445
      Width           =   7125
   End
   Begin VB.Label Label3 
      Caption         =   "Click Here To Cancel At Any Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   945
      Width           =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "Files..."
      Height          =   195
      Left            =   3735
      TabIndex        =   10
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Folders..."
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   1350
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A simple directory recursion example using pure vb code
' You can modify it however you wish, All file operationms now supports
' file mask list in this format "*.ext;*.ext;*.ext" an so one where ext is the file extention to mask
' By Lewis Miller

Private Sub Command1_Click()
Dim Files As New Collection, varFile As Variant
 
 Cancelled = False
 List1.Clear
 List2.Clear
 
 If Check1 Then
   RecurseFiles Files, Text1, InputBox("Enter file mask to use separated by the ; symbol or use default", "File Mask", "*.*") 'gets all files including files in sub folders
 Else
   GetFiles Files, Text1    'gets all files Not including sub folders
 End If
 
 For Each varFile In Files
    List2.AddItem varFile
 Next
 
End Sub

Private Sub Command2_Click()
Dim Folders As New Collection, varFolder As Variant
  
 Cancelled = False
 List1.Clear
 List2.Clear
   
   If Check1 Then
     RecurseFolders Folders, Text1 'gets all folders including subfolders
   Else
     GetFolders Folders, Text1     'gets all subfolders in directory but not the folders in those
   End If
   
   For Each varFolder In Folders
      List1.AddItem varFolder
   Next
 
End Sub

Private Sub Command3_Click()
Dim DirList As New Collection, Files As New Collection, varItem As Variant
  
 Cancelled = False
 List1.Clear
 List2.Clear
 lblStatus = "Retrieving Directory Information..."
 
 If Check2 Then
   RecurseAll DirList, Text1 'gets all folders and files including subs in 1 collection
 Else
   RecurseSeperate Files, DirList, Text1 'gets all folders and files including sub\ but files and folders are seperate
   lblStatus = "Found " & CStr(Files.Count) & " Folders, Containing " & CStr(DirList.Count) & " Files... Adding to List."
   DoEvents
     For Each varItem In Files
        List2.AddItem varItem
     Next
 End If
  
    For Each varItem In DirList
       List1.AddItem varItem
    Next
   
   lblStatus = "Search Complete."

End Sub


Private Sub Command4_Click()

  Dim Foldersize As Double
 
   lblStatus = "Please Wait, Calculating Folder Size..."
 
   Foldersize = GetFolderSize(Text1.Text, Check1) 'calculate all files in folder
 
   MsgBox "Folder Size: " & CStr(Foldersize) & " Bytes."
 
   lblStatus = "Folder Size Check Complete."
   
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim X As Long, File() As FileInformation

 Cancelled = False
 List1.Clear
 List2.Clear
 lblStatus = "Retrieving Directory Information..."
 
'call the sub to retrieve all files including subfolders ... yay!
DirFileInformation File, Text1, Check1, "*.bmp;*.gif;*.jpg"

'test to see if any files were returned (if error then no files)
If IsError(UBound(File)) = False Then
   
   For X = 0 To UBound(File) 'loop thru all an display
     List1.AddItem File(X).Folder
     List2.AddItem File(X).Title & " : " & CStr(File(X).Size)
     If (X Mod 10) = 0 Then DoEvents
   Next X
   
   lblStatus = "Search Complete... Found " & CStr(UBound(File)) & " Files."
Else

   lblStatus = "Search Complete... Found 0 Files."

End If

End Sub

Private Sub Form_Load()
  Cancelled = False 'forces the module to load
                    'so the first time isnt as slow :)
End Sub

Private Sub Label3_Click()
 Cancelled = True 'cancel Dir() searching
End Sub
