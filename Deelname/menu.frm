VERSION 5.00
Begin VB.Form MenuVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "menu.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   18.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   24.125
   Begin VB.CommandButton VerwijderDeelnamesVoorPeriodeKnop 
      Caption         =   "&Verwijder deelnames voor periode."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton ToonDeelnamesVoorPeriodeKnop 
      Caption         =   "&Toon deelnames voor periode."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton VerwijderDeelnamesKnop 
      Caption         =   "&Verwijder deelnames."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton EindeKnop 
      Cancel          =   -1  'True
      Caption         =   "&Einde."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton ToonDeelnamesVoorDatumKnop 
      Caption         =   "&Toon deelnames voor datum."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton ToonDeelnamesVoorLidnummerKnop 
      Caption         =   "&Toon deelnames voor lidnummer."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton VoerDeelnamesInKnop 
      Caption         =   "&Voer deelnames in."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "MenuVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het menu venster.
Option Explicit



'Deze procedure sluit dit programma af.
Private Sub EindeKnop_Click()
On Error GoTo Fout
   Unload MenuVenster
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
   MenuVenster.Left = 128
   MenuVenster.Top = 128
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure opent het venster met de deelname datums.
Private Sub ToonDeelnamesVoorDatumKnop_Click()
On Error GoTo Fout
   DeelnameDatumsVenster.Show vbModal
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



'Deze procedure geeft opdracht om de deelnames voor de opgegeven periode te tonen.
Private Sub ToonDeelnamesVoorPeriodeKnop_Click()
On Error GoTo Fout
Dim TotDatum As String
Dim VanDatum As String

   VraagPeriode VanDatum, TotDatum
   
   DeelnameVenster.VanDatum = VanDatum
   DeelnameVenster.TotDatum = TotDatum
   
   If Not (VanDatum = vbNullString Or TotDatum = vbNullString) Then
      DeelnameVenster.Show
      DeelnameVenster.SetFocus
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft opdracht om het aantal deelnames van het opgegeven lidnummer voor de opgegeven datum te tonen.
Private Sub ToonDeelnamesVoorLidnummerKnop_Click()
On Error GoTo Fout
   ToonLidnummerDeelname
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure geeft de opdracht om te beginnen met de invoer van deelnames.
Private Sub VoerDeelnamesInKnop_Click()
On Error GoTo Fout
   VoerDeelnamesIn
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



'Deze procedure start de deelname bewerkings functie in verwijder modus.
Private Sub VerwijderDeelnamesKnop_Click()
On Error GoTo Fout
   VerwijderDeelnames
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure geeft opdracht om de deelnames binnen de opgegeven periode te verwijderen.
Private Sub VerwijderDeelnamesVoorPeriodeKnop_Click()
On Error GoTo Fout
   VerwijderDeelnamesVoorPeriode
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


