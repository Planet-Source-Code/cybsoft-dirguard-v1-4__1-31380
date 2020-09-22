VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmGuard 
   Caption         =   "Directory Guard"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "frmGuard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFileInfo 
      Height          =   2010
      Left            =   3960
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.FileListBox HiddenFilelist 
      Height          =   2040
      Left            =   1680
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   5280
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.Frame Frame7 
         Caption         =   " Logfile "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4215
         Left            =   5640
         TabIndex        =   17
         Top             =   120
         Width           =   5055
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   255
            Left            =   1800
            TabIndex        =   28
            Top             =   3840
            Width           =   495
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Exit"
            Height          =   255
            Left            =   3960
            TabIndex        =   24
            Top             =   3840
            Width           =   975
         End
         Begin VB.CommandButton cmdClearLoG 
            Caption         =   "Clear"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   3840
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   1080
            TabIndex        =   22
            Top             =   3840
            Width           =   495
         End
         Begin RichTextLib.RichTextBox rtbChangedfiles 
            Height          =   3495
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   6165
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"frmGuard.frx":030A
         End
         Begin VB.Label Label5 
            Caption         =   "Changes"
            Height          =   255
            Left            =   2880
            TabIndex        =   27
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label lblChanges 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   2400
            TabIndex        =   26
            Top             =   3840
            Width           =   375
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Guarding "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   2880
         TabIndex        =   16
         Top             =   3600
         Width           =   2655
         Begin VB.CommandButton cmdGuardStop 
            Caption         =   "&Stop"
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "&Start"
            Height          =   360
            Left            =   240
            TabIndex        =   20
            Top             =   260
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Result & Status "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   2880
         TabIndex        =   11
         Top             =   2280
         Width           =   2655
         Begin VB.TextBox txtChanges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "No changes"
            Top             =   800
            Width           =   1695
         End
         Begin VB.TextBox txtStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Idle"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Change"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   820
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Status"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   380
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Settings "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   2880
         TabIndex        =   6
         Top             =   1080
         Width           =   2655
         Begin VB.TextBox txtTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "5 Sec"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtFiles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Update time"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Number of files"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Refresh time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   2880
         TabIndex        =   4
         Top             =   120
         Width           =   2655
         Begin MSComctlLib.Slider Slider1 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   10
            SmallChange     =   5
            Min             =   5
            Max             =   50
            SelStart        =   5
            Value           =   5
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Navigation "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         Begin VB.FileListBox lstFiles 
            Appearance      =   0  'Flat
            Height          =   1395
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   2640
            Width           =   2415
         End
         Begin VB.DirListBox lstMap 
            Appearance      =   0  'Flat
            Height          =   1665
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   2415
         End
         Begin VB.DriveListBox Drivestation 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "frmGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClearLoG_Click()
rtbChangedfiles.Text = ""
lblChanges.Caption = ""
cmdStart.SetFocus           ' the buttons are so small, the caption is deformed
End Sub                     ' seting the focus on another button resolves this

Private Sub cmdPrint_Click()
    rtbChangedfiles.SelPrint (Printer.hDC)  ' could be better but it works
    Printer.EndDoc                          ' just print the lof and eject page
    cmdStart.SetFocus                       ' Buttoncaption restored
End Sub

Private Sub Form_Activate()
'I you want to show a programversion use the folowing :
Me.Caption = " Directory Guard V" & App.Major & "." & App.Minor
'these values can be entered or automaticly set by rightclicking on the project
'nouw look in the programproperty on the tab named MAKE
'If you will you can add the revision by adding : & "." & App.Revision

txtFiles.Text = lstFiles.ListCount          'we need some info on the form
End Sub

Private Sub Drivestation_Change()           ' we change drive
Call StopTimer                              ' FIRST stop the timer
    On Error GoTo error                     ' in case a disk is not available
    lstMap.Path = Drivestation.Drive        ' set the directory for the map-list
    Exit Sub
error:                                      ' disk was not available
    Dim answer As Integer
    answer = MsgBox(Err.Description, 5, "Device error !")
    If Annswer = 4 Then Resume              ' they pressed ok
End Sub

Private Sub lstMap_Change()                 ' the map-information
Call StopTimer                              ' if not stopped...stop it now
lstMap.Refresh                              ' refresh it
lstFiles.Path = lstMap.Path                 ' set the map for the filelisting
HiddenFilelist.Path = lstMap.Path           ' needed to detect what file is moved
lstFileInfo.Clear                           ' clear stored filedates /time
txtFiles.Text = lstFiles.ListCount          ' put number of files on the form
End Sub
Private Sub cmdStart_Click()                ' start timer by changing value from 0 to
txtStatus.Text = "Guard started"            ' what are we doing
txtTime.Text = Slider1.Value & " Sec"       ' Adjust value on form
Timer1.Interval = Slider1.Value * 1000      ' the delaytime
lstFiles.Refresh                            ' start fresh from now, the box is already
HiddenFilelist.Refresh                      ' filled so you start logging before you start
Call FileCheckDateTime                      ' store files date and time
End Sub
Private Sub cmdGuardStop_Click()
Call StopTimer                              ' make timer interval = 0
txtStatus.Text = "Guard stopped"            ' for informaton
End Sub
Private Sub Slider1_Change()                ' change the update-time for the timer
Dim updatetime As Integer                   ' there is no need for all declarations
                                            ' its best to get used to it and use them
updatetime = Slider1.Value                  ' whats the value of the slider ?

If Timer1.Interval <> 0 Then                ' unless the timer is stopped . . .  !!
    Timer1.Interval = (updatetime * 1000)   ' make seconds from milliseconds
    txtTime = CStr(updatetime) & " Sec"     ' put it in a string
End If
' what can be done is "dynamic adjustment" of the delay-time by taking the number
' of files in the Filelistbox (lstFiles) using lstFiles.listcount.
' Now devide this number by 2 so 10 files would give you value 5.
' Multiply this with 1000 and you have a Timer1.interval= 5000 (5 Sec)
' If you have 100 files this would mean a interval of 50000 (50 sec's)
' this can be to long so you can also use a fixed delaytime calculated from
' the number of files and use :
' If lstFiles.ListCount > 100 Then Timer1.Interval = 7000
' If lstFiles.ListCount > 200 Then Timer1.Interval = 10000 (and so on)

End Sub

Private Sub Timer1_Timer()
Dim newfiles As Integer
Dim buffer As String
Dim ReadCounter As Integer
Dim CurrentItem As String

txtStatus.Text = "Updating"                     ' what are we doing
txtStatus.Refresh                               ' with short listings it would not be visible
lstFiles.Refresh                                ' obtain new data
newfiles = lstFiles.ListCount                   ' how many files are there NOW ?
                                                ' old value stored in txtFiles.text
                                                
    If newfiles = Val(txtFiles.Text) Then       ' nothing has changed
        Call CheckThatFile                      ' check for alterations
        txtStatus.Text = "Await update"         ' tell what we are doing
        Exit Sub                                ' get out
    End If

    If newfiles > Val(txtFiles.Text) Then       ' something is added
        txtChanges.Text = CStr(newfiles - Val(txtFiles.Text)) & " File(s) added"
       
'We should have a difference between the content of HiddenFilelist and lstFiles
'That difference is the file moved or added, this we need to find.
'This can be done in several way's but i use "instr" not the pretty but
'suitable for this example.
'But first i need to read the content of my hiddenFileList into a
'texststring, lets call it buffer
       
    For ReadCounter = 0 To HiddenFilelist.ListCount     ' zero is the first item not 1 (!)
        CurrentItem = HiddenFilelist.List(ReadCounter)
        buffer = buffer & CurrentItem & Space(1)        ' all filenames devided by space
    Next ReadCounter
        
'now we only have to read the files in the lstFiles and see what's new there

    For ReadCounter = 0 To lstFiles.ListCount
        CurrentItem = lstFiles.List(ReadCounter)
            If InStr(buffer, CurrentItem) = 0 Then  ' its a new file
                rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  " & CurrentItem & "  " & "was Added" & vbCrLf
                lblChanges.Caption = Val(lblChanges.Caption) + 1    'number of entrys in log
                
                lstFiles.Refresh                ' update the lists for change detection
                HiddenFilelist.Refresh          ' If we do not, a new file is seen as eddited
                lstFileInfo.Clear               ' or the sync between the filenamelist
                Call FileCheckDateTime          ' and the date-timelist is gone
             End If                             ' then all are seen as changed
    Next ReadCounter
    End If

'now the same but the other way around

    If newfiles < Val(txtFiles.Text) Then       ' somthing has been moved/deleted
        txtChanges.Text = CStr(Val(txtFiles.Text) - newfiles) & " File(s) removed"
        
    For ReadCounter = 0 To lstFiles.ListCount
        CurrentItem = lstFiles.List(ReadCounter)
        buffer = buffer & CurrentItem & Space(1)  'all filenames devided by space
    Next ReadCounter
        
    For ReadCounter = 0 To HiddenFilelist.ListCount
        CurrentItem = HiddenFilelist.List(ReadCounter)
            If InStr(buffer, CurrentItem) = 0 Then  'its a new file
                rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  " & CurrentItem & "  " & "was Removed" & vbCrLf
                lblChanges.Caption = Val(lblChanges.Caption) + 1
                lstFiles.Refresh
                lstFileInfo.Clear
                HiddenFilelist.Refresh
                Call FileCheckDateTime
             End If
    Next ReadCounter
        
    End If
    
HiddenFilelist.Refresh                          ' don't fotget or get 1 change only
txtFiles.Text = CStr(newfiles)                  ' update our filecount-value
txtStatus.Text = "Await update"                 ' update our tekst
Me.Refresh                                      ' refresh the lot
End Sub
Private Sub FileCheckDateTime()
Dim Filepath As String
Dim ReadCounter As Integer
Dim CurrentItem As String
Dim Needed As String
' We now need to store the file-date and time
' To do this we will read the data / items from HiddenFilelist.list and store
' them in a hidden listbox in the same order as the files are (in sync).

' First we need to see if the itemscount matches, if so, dont add again

If HiddenFilelist.ListCount = lstFileInfo.ListCount Then    ' they match
    Exit Sub
End If
' Then we need the path where the files monitored will be

Filepath = HiddenFilelist.Path & "\"

For ReadCounter = 0 To HiddenFilelist.ListCount
    CurrentItem = HiddenFilelist.List(ReadCounter)  ' contains the filename
        If CurrentItem = "" Then GoTo Ready         ' we are ready
     Needed = FileDateTime(Filepath & CurrentItem)
        lstFileInfo.AddItem (Needed)
Next ReadCounter
Ready:
End Sub
Private Sub CheckThatFile()
Dim Filepath As String
Dim ReadCounter As Integer
Dim CurrentItem As String
Dim Needed As String
Dim Stored As String

' Now we need to see if the file-dates have changed (or the filesize !!)
' Therefore we are reading the items in the lstFiles (FileListbox)
' Then we check if the time and date still matches the stored one.
' We only do this test if no files are deleted or added !!.
' The data in both boxes should be the same, also the item-order (in Sync).

Filepath = HiddenFilelist.Path & "\"        'where the monitored files are

For ReadCounter = 0 To lstFiles.ListCount
    CurrentItem = lstFiles.List(ReadCounter)
        If CurrentItem = "" Then GoTo Ready                     ' we are ready
            Needed = Trim(FileDateTime(Filepath & CurrentItem)) ' be safe and trim
            Stored = Trim(lstFileInfo.List(ReadCounter))        ' the stored fileinfo
                                                                ' compair the 2 strings
            If Stored <> Needed Then                            ' we have difference !
                rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  " & CurrentItem & "  " & " has been changed" & vbCrLf
                lblChanges.Caption = Val(lblChanges.Caption) + 1

            End If
Next ReadCounter                ' more files altered (in a network-enviroment possibe)

Ready:
lstFiles.Refresh                ' we need to update all lists / data
lstFileInfo.Clear               ' this one must be cleared or all info is appended
HiddenFilelist.Refresh          ' refresh to read the new data from
Call FileCheckDateTime          ' re-read / store the files again
End Sub

Private Sub StopTimer()     ' this sub is seperate because its used on several occasions
Timer1.Interval = 0         ' Making the interval = 0 stops the timer
txtStatus.Text = "Idle"     ' update the status
Me.Refresh                  ' and refresh the form
End Sub
Private Sub cmdSave_Click()
frmSave.Show                ' show the form containing a dirty savefile
End Sub

Private Sub cmdExit_Click()
Unload frmSave
Unload frmGuard
End
End Sub
