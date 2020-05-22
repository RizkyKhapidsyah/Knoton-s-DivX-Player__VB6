VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Help Hotkeys"
   ClientHeight    =   2220
   ClientLeft      =   8625
   ClientTop       =   5700
   ClientWidth     =   3075
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   3075
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Show
Print "Developer Kenneth Hedman (Knoton)"
Print "http://www.knoton.dns2go.com" & vbCrLf
Print "The HotKeys are:" & vbCrLf & _
        "CTRL + F5 Play/Pause" & vbCrLf & _
        "CTRL + F6 Stop" & vbCrLf & _
        "CTRL + F7 Fullscreen On/Off" & vbCrLf & _
        "CTRL + Add key Increase Moviesize" & vbCrLf & _
        "CTRL + Subtract key decrease Moviesize" & vbCrLf & _
        "CTRL + Down Arrow decrease volume" & vbCrLf & _
        "CTRL + Up Arrow increase volume"
End Sub
