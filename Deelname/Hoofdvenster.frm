VERSION 5.00
Begin VB.MDIForm HoofdVenster 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Deelname Teller"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7830
   Icon            =   "Hoofdvenster.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu InformatieMenu 
      Caption         =   "&Informatie"
   End
End
Attribute VB_Name = "HoofdVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het hoofdvenster.
Option Explicit

'Deze procedure geeft opdracht om informatie over dit programma te tonen.
Private Sub InformatieMenu_Click()
On Error GoTo Fout
   ToonProgrammainformatie
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub MDIForm_Load()
On Error GoTo Fout
   HoofdVenster.WindowState = vbMaximized

   HoofdVenster.Width = Screen.Width / 1.5
   HoofdVenster.Height = Screen.Height / 1.5
   
   MenuVenster.Show
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


