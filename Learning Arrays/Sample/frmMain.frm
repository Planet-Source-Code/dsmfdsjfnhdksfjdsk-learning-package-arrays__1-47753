VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Menus at Runtime"
   ClientHeight    =   1575
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Menu"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame fmeMain 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0000
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "MainMenu"
      Begin VB.Menu mnuMainSubItem 
         Caption         =   "SubItem"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MenuCount As Integer ' We need to know how many menus have been loaded already

Private Sub cmdAdd_Click()
    MenuCount = MenuCount + 1                                  ' Just add one to the MenuCount so we don't load a menu that already loaded
    Load mnuMainSubItem(MenuCount)                             ' Load an instance of mnuMainSubItem
    mnuMainSubItem(MenuCount).Caption = "Instance of SubItem"  ' Give the new instance a caption
End Sub

Private Sub mnuMainSubItem_Click(Index As Integer)                                              ' The index is there so we know which menu has been clicked
    Select Case Index                                                                           ' Depending on the Index, we will tell VB what to do.
        Case 0                                                                                  ' If the case is zero
            MsgBox "This is the first menu.", vbInformation, "MenuInstance"                     ' Tell VB what to do
        Case 1                                                                                  ' If the case is 1
            MsgBox "This is an instance of the first menu.", vbInformation, "MenuInstance"      ' Tell VB what to do
        Case Else                                                                               ' If the case is something else
            MsgBox "Just another instance of the first menu...", vbInformation, "MenuInstance"  ' Tell VB what to do.
    End Select                                                                                  ' End our set of instructions
End Sub
