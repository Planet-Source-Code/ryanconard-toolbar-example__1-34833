VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolBar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ToolBar / ImageList Example by Ryan Conard"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmToolBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "Type ""exit"" here to quit!"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      Begin VB.Label Label1 
         Caption         =   $"frmToolBar.frx":0442
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
   End
   Begin MSComctlLib.Toolbar tlbNew 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   1535
      ButtonWidth     =   1614
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Controls"
            Key             =   "Controls"
            Description     =   "ToolBar Example by Ryan Conard"
            Object.ToolTipText     =   "Button Example by Ryan Conard"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Internet"
            Key             =   "Internet"
            Description     =   "ToolBar Example by Ryan Conard"
            Object.ToolTipText     =   "ToolBar Example by Ryan Conard"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Floppy"
            Key             =   "Floppy"
            Description     =   "ToolBar Example by Ryan Conard"
            Object.ToolTipText     =   "ToolBar Example by Ryan Conard"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Disk Drives"
            Key             =   "Disk Drives"
            Description     =   "ToolBar Example by Ryan Conard"
            Object.ToolTipText     =   "ToolBar Example by Ryan Conard"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Information"
            Key             =   "Information"
            Description     =   "ToolBar Example by Ryan Conard"
            Object.ToolTipText     =   "ToolBar Example by Ryan Conard"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":0541
            Key             =   "Controls"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":0995
            Key             =   "Internet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":0DE9
            Key             =   "Floppy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":123D
            Key             =   "Disk Drives"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBar.frx":1691
            Key             =   "Information"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The OSO [Open-Source Only] Team
'Code by Ryan Conard
'rconard@localnet.com


Private Sub Text1_Change()

'This is my own thing I decided to put in..
'It's stupid, but if you like it.. you can use it..
'Thanks!
Dim intExit As Integer
Select Case Text1
Case Is = "exit"
Unload Me
MsgBox "Bye.. thanks for using my app.", , "Bye.."
End
End Select

End Sub

'What to do if this code doesn't work:
'1: Click the "Project" menu
'2: Click the "Components" button
'3: Scroll about 3/4 of the way down and check "Microsoft Windows Common Controls 6.0"
'4: Click "Ok" and your done!
'
'If for any reason this doesn't work afterwards, please email me..
'My address is listed below in the code somewhere..
'Enjoy!

Private Sub tlbNew_ButtonClick(ByVal Button As MSComctlLib.Button)

Rem: Respond to button clicks.. Integers are your friends!
Dim msgPress As Integer

'Display a MsgBox depending
'on which ToolBar button the user clicked
Select Case Button.Key 'The "Button.Key" text call upon the ToolBar buttons KEY..
                       'access this by right-clicking the ToolBar and selecting
                       'properties.
                       
Case Is = "Controls":
msgPress = MsgBox("You pressed the 'Controls' button!", , "ToolBar Example")
Case Is = "Internet":
msgPress = MsgBox("You pressed the 'Internet' button!", , "ToolBar Example")
Case Is = "Floppy":
msgPress = MsgBox("You pressed the 'Floppy' button!", , "ToolBar Example")
Case Is = "Disk Drives":
msgPress = MsgBox("You pressed the 'Disk Drives' button!", , "ToolBar Example")
Case Is = "Information":
msgPress = MsgBox("You pressed the 'Information' button! -This code was created by rconard@localnet.com!", , "ToolBar Example")

'The code is finished!
End Select
End Sub
