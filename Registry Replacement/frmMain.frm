VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "reg replacement"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1425
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   1425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Get Setting"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Setting"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem //
Rem // The 'modReg_Rep' module will replace the Visual Basic 'SaveSetting' and 'GetSetting' subs,
Rem // to these custom ones... meaning that you will now be able to save all your visual basic
Rem // Projects to a custom location in the registry... (I use a 'Local_Machine' location that way
Rem // my settings are global.)
Rem //
Rem // This project was mainly written so i could save to my own folder in the registry... not the
Rem // Dodgy Microsoft Visual Basic folder... making my projects more professional.
Rem //
Rem // you can customize the location of where the keys are saved, and read from under the
Rem // 'GetSetting' and 'SaveSetting' subs.
Rem //
Rem //
Rem // Code by David Nedved. 2005
Rem // Website: www.datosoftware.com
Rem //

Private Sub Command1_Click()
Rem // Save a key to the registry (in the Main LOCAL MACHINE hkey not the LOCAL SETTINGS hkey)
Dim msgStr As String
msgStr = InputBox("Please type in what you want to save to the Registry.", "Save Setting into Registry?", "This is a Test Save... 'info will be saved to {HKEY_LOCAL_MACHINE\Software\DaTo Software\RegTest\Section As String\Key As String}'")
SaveSetting "RegTest", "Section As String", "Key As String", msgStr
End Sub

Private Sub Command2_Click()
Rem // Get the above message that is saved to the registry...
Rem // to see that this realy works try saving a key... then closing the program and opening it again...
Rem // It will REALY get the key from the registry.
MsgBox GetSetting("RegTest", "Section As String", "Key As String", "Hmm no info could be found in the Registry... please Save a Registry Key first, as this is only a Default Message"), vbInformation, "hello"
End Sub
