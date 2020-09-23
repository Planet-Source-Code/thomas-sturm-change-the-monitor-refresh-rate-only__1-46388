VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Monitor Refresh Rate - Test (c) Thomas Sturm 2003"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraGood 
      BorderStyle     =   0  'Kein
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   6135
      Begin VB.CommandButton cmdEnd 
         Caption         =   "End"
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Shape shpGood 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Ausgefüllt
         Height          =   255
         Left            =   0
         Shape           =   3  'Kreis
         Top             =   285
         Width           =   735
      End
      Begin VB.Label lblDummy 
         Alignment       =   2  'Zentriert
         Caption         =   "Your Frequency-Settings are optimal."
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   315
         Width           =   3735
      End
   End
   Begin VB.Frame fraBad 
      BorderStyle     =   0  'Kein
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   6135
      Begin VB.CommandButton cmd85Hz 
         Caption         =   "Set Frequency to 85 Hz"
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblDummy 
         Caption         =   "Your Frequency-Settings are not optimal, you should see a slight flickering."
         Height          =   410
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   5055
      End
      Begin VB.Shape shpBad 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   255
         Left            =   0
         Shape           =   3  'Kreis
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.Line linDummy 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   6000
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Line linDummy 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   120
      X2              =   6000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblDummy 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Your Monitor runs at"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblHz 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "%lblHz%"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   I wrote this program for the company I work for, we recently switched from NT4 to Win2000
'   (WHO LAUGHED ?!?!?!), and since then many users mentioned that their monitor began to
'   "flicker" ( I work as a PC-Technician), but 99% of these problems were because our
'   Win2000-Migration Team ALWAYS leaves the Refresh-Rate at 60Hz. Me and my collegues didn't
'   always want to have to see the user personally for that reason, so I wrote this proggy,
'   which sets the Frequency to 85 Hz (this is the max. that our VGA-Cards can handle at
'   1280x1024x16) (WHO LAUGHED AGAIN ?!?!?!?!). Also, I never found a prog on PSC, which lets
'   you adjust JUST the Frequency, only the Resolution and the Color-Depth.

Option Explicit

'-----Constants
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000          'You would use these to adjust the
Private Const DM_PELSWIDTH = &H80000           'Screen-Height, Width and Color-Depth,
Private Const DM_PELSHEIGHT = &H100000         'I just needed the Frequency !
'--------------------
Private Const DM_DISPLAYFREQUENCY = &H400000    'Never saw this one in ANY submission
'--------------------
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const ENUM_CURRENT_SETTINGS = &HFFFF - 1

'-----TYPE-Structure
Private Type DEVMODE
   dmDeviceName As String * CCDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

'-----APIs
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Private Sub cmd85Hz_Click()
'This holds the structure of all settings
Dim DevM As DEVMODE
'What the name says, was the change successful ?
Dim result As Long
'Get the current settings
Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, DevM)
'Only change the frequency
DevM.dmFields = DM_DISPLAYFREQUENCY
'Tell Windows what frequency you want (here 85 Hz)
DevM.dmDisplayFrequency = 85
'First, TEST the new settings only
result = ChangeDisplaySettings(DevM, CDS_TEST)
'What did the function return ?
Select Case result
    Case DISP_CHANGE_RESTART
         'The function said that you have to restart.
         'You should never end up here, since we only change the frequency
         MsgBox "You have to restart for the changes to take effect.", vbYesNo + vbSystemModal, "Info"
    Case DISP_CHANGE_SUCCESSFUL
        'Usually, you end up here, the function said OK
        'Update the Registry with the new settings, so you keep it after the next restart.
        result = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
        MsgBox "Frequency successfully changed !", vbInformation + vbSystemModal + vbOKOnly, "Änderung erfolgreich"
End Select
End Sub

Private Sub cmdEnd_Click()
'Hmmm, what does this do ...
Unload Me
End Sub

Private Sub Form_Load()
'This holds the structure of all settings
Dim DevM As DEVMODE
'Get the current settings
Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, DevM)
'Show the current refresh-rate
lblHz.Caption = DevM.dmDisplayFrequency & " Hz"
'Prepare the GUI (or whatever you want to call it...)
'If the current frequency is below 85 Hz ...
If DevM.dmDisplayFrequency < 85 Then
    'Hide one frame and show the other
    fraBad.Visible = True
    fraGood.Visible = False
'Frequency is equal to or bigger than 85 Hz
ElseIf DevM.dmDisplayFrequency >= 85 Then
    'Hide one frame and show the other (I know that I am somewhat repetitive ...)
    fraBad.Visible = False
    fraGood.Visible = True
End If
End Sub
