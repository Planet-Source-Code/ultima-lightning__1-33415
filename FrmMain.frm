VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lighning"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timing 
      Interval        =   75
      Left            =   3480
      Top             =   3240
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   9
      Left            =   0
      Picture         =   "FrmMain.frx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   8
      Left            =   0
      Picture         =   "FrmMain.frx":3034E
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   7
      Left            =   0
      Picture         =   "FrmMain.frx":60392
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   6
      Left            =   0
      Picture         =   "FrmMain.frx":903D6
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   5
      Left            =   0
      Picture         =   "FrmMain.frx":C041A
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   4
      Left            =   0
      Picture         =   "FrmMain.frx":F045E
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   3
      Left            =   0
      Picture         =   "FrmMain.frx":1204A2
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   2
      Left            =   0
      Picture         =   "FrmMain.frx":1504E6
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   1
      Left            =   0
      Picture         =   "FrmMain.frx":18052A
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Image Lightning 
      Height          =   3840
      Index           =   0
      Left            =   0
      Picture         =   "FrmMain.frx":1B056E
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************'
'* By Michael Kissner (<<ULTIMA>>) *'
'*                                 *'
'* This program randomly chooses a *'
'* frame of a Lightning and then   *'
'* displays it on the form.        *'
'***********************************'

Option Explicit

Dim LastPic As Integer              'A Variable that shows what
                                    'Pic to draw next

Private Sub Form_Load()             'When the Form is loaded

On Error Resume Next                'The error handler, incase
                                    'anything happens
    Timing.Interval = 75            'Sets the interval to 75/100
                                    'secodns. (That means that
End Sub                             '1.5 frames are shown each
                                    'second
Private Sub Timing_Timer()

On Error GoTo ErrorHandler          'The ErrorHandler, no realy
                                    'needed here, but you never know.
    Me.Picture = Lightning(LastPic).Picture 'Draws the Picture
                                    'on the main Form
    LastPic = LastPic + (Rnd * 5)   'Gets a new Value for LastPic
    LastPic = LastPic - (LastPic Mod 1) 'Removes the numbers
                                    'after the dot.
    If LastPic >= 9 Then            'This makes sure that the
        LastPic = 0                 'number is not bigger then 9.
    End If
    
    Exit Sub                        'Exits sub, so the error msg
                                    'won't be displayed.
ErrorHandler:                       'The program executes this
    If MsgBox("An error has occured, the program will now shutdown" _
                , vbOKOnly, "Error") = vbOK Then 'when an error accures.
        End                         'When the user presses ok
    End If                          'the program exits.
    
End Sub                             'Guess ;)
