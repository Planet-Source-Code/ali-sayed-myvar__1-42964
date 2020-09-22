VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyVar test"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTxt 
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3625
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   1e7
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"FrmMain.frx":0000
   End
   Begin VB.CommandButton CmdMyVar 
      Caption         =   "MyVar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton CmdVar 
      Caption         =   "VB string"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
' Project:  MyVar   v.1.0                   *
' Author:   Ali Sayed                       *
' E-Mail:   AliSayed_7@Yahoo.com            *
' Date:     5/02/2003                      *
' Copyright © 2003 Ali Sayed                *
' Please let me know if you like it.        *
' For more information mail me.             *
'********************************************
Option Explicit
'The MyVar Class is for using instead of String var
'for handling big string var for speed issue

Private Sub CmdMyVar_Click()
Dim MyVar As New ClsMyVar
Dim i As Long
' reset MyVar text
MyVar.Clear

'test part
'with MyVar you can see the progress inedcator keeps
' fast till the end of loop
For i = 1 To 1000
'note In MyVar you dont use the statment
'(MyVar = MyVar & newtext)
'cause MyVar already add newtext to its old value
    MyVar.Text = String$(1000, "A") & vbCrLf
    ProgressBar2.Value = i
Next

'load vars into rich text
RichTxt.Text = MyVar.Text
RichTxt.Refresh
'make sure they as the same
MsgBox Len(RichTxt.Text)
End Sub

Private Sub CmdVar_Click()
Dim StrVar As String
Dim i As Long

'test part
'with StrVar you can see the progress inedcator go
'slowley as it go foreward
For i = 1 To 1000
    StrVar = StrVar & String$(1000, "A") & vbCrLf
    ProgressBar1.Value = i
Next

'load vars into rich text
RichTxt.Text = StrVar
RichTxt.Refresh
'make sure they as the same
MsgBox Len(RichTxt.Text)
End Sub
