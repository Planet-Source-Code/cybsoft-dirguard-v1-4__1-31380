VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save File"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3510
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdDoIt 
         Caption         =   "Save"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Filename :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmSave.Hide                                            ' Hide this form
frmGuard.Refresh                                        ' refresh the mainform
frmGuard.cmdStart.SetFocus                              ' Setting the focus
End Sub

Private Sub cmdDoIt_Click()

' we only want to save in the current path so common-dialog is not needed
' for this action just write the file where the program is.
' You can add commondialogs , add a fixed path / directory and so on
' but i want this to be simple.

textbuffer = Trim(frmGuard.rtbChangedfiles.Text)        ' here is the logdata

Looping:                                                ' here i use a label

Newfile = Trim(txtFilename.Text)                        ' here is the filename

                                                        ' lets check some things:

If Len(textbuffer) = 0 Then                             ' the log is empty
    frmSave.Hide
    frmGuard.txtStatus.Text = "Nothing to save!"        ' put it on mainform
    Exit Sub
End If

If Len(Newfile) = 0 Then                                ' no filename
    txtFilename.SetFocus
    GoTo Looping                                        ' not nice but effective
End If

If Right(Newfile, 4) <> ".txt" Then                     ' we want to make txtFles
    Newfile = Newfile & ".txt"                          ' if not we add extention
End If

On Error GoTo error                                     ' works, not eligant
Me.MousePointer = vbHourglass                           ' why not...change the mouse

Open Newfile For Output As #1                           ' Open a file number #1
Print #1, textbuffer                                    ' print to this file
Close #1                                                ' close it
frmSave.Hide                                            ' hide this form
frmGuard.Refresh                                        ' refresh the main form
Me.MousePointer = vbNormal                              ' and change mouse back
Exit Sub

error:
X = MsgBox("Something went wrong while saving!", vbOKOnly, "Error")
frmSave.Hide
frmGuard.Refresh
End Sub
