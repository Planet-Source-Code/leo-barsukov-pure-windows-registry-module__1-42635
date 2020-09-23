VERSION 5.00
Begin VB.Form RegTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registry Module Testing (Purely Windows - No Api)"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Roll Back"
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox ProgName 
      Height          =   285
      Left            =   3120
      TabIndex        =   16
      Text            =   "TestProg"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Associate"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox ico 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Text            =   "yapp.exe,0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox app 
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Text            =   "yapp.exe"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox ext 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "rex"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Value 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Text            =   "Value"
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Folder 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "NewKey\SubKey\Setting"
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox RootFolder 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label7 
      Caption         =   "TestProg"
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Icon:"
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Open With:"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Ext."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Value"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Folder"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Root Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "RegTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Step 1:
'Dim the Regstry Class Module
Dim Reg As New Registry

'''''''''''''''''''''''
'Step 2:
'''''''''''''''''''''''
Private Sub Command1_Click()

'--/Write to Registry     *RootFolder *Folder *Value
    Reg.WriteKey RootFolder.ListIndex, Folder, Value

End Sub

Private Sub Command2_Click()

'--/Read From Registry            RootFolder *Folder
    Value = Reg.ReadKey(RootFolder.ListIndex, Folder)

End Sub

Private Sub Command3_Click()

'--/Delete From Registry   *RootFolder *Folder
    Reg.DeleteKey RootFolder.ListIndex, Folder

End Sub

Private Sub Command4_Click()
    
    Reg.AssociateFile ext.Text, app.Text, ProgName.Text, ico.Text

End Sub

Private Sub Command5_Click()
Reg.RollBackFile ext, ProgName
End Sub

Private Sub ext_Change()
For i = 1 To Len(ext)
If Mid(ext, i, 1) = "." Then
ext.SelStart = i - 1
ext.SelLength = 1
ext.SelText = ""
ext.SelStart = Len(ext)
End If
Next
End Sub

Private Sub Form_Load()
    
    'Reset RootFolder ListIndex
    RootFolder.ListIndex = 0

End Sub
