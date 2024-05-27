VERSION 5.00
Begin VB.Form DeelnameDatumsVenster 
   Caption         =   "Deelname Datums"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   ClipControls    =   0   'False
   Icon            =   "datums.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   31.125
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   600
      ScaleHeight     =   1.563
      ScaleMode       =   4  'Character
      ScaleWidth      =   25.125
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
      Begin VB.CommandButton AnnulerenKnop 
         Cancel          =   -1  'True
         Caption         =   "&Annuleren"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton ToonDeelnamesKnop 
         Caption         =   "&Toon Deelnames"
         Default         =   -1  'True
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.ListBox DatumLijst 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "DeelnameDatumsVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het venster waarin alle deelname datums worden getoond.
Option Explicit

'Deze procedure sorteert de getoonde deelname datums.
Private Sub SorteerLijst()
On Error GoTo Fout
Dim AndereDatum As Long
Dim AnderePeriode As Long
Dim Datum As Long
Dim DMJ As Long
Dim Periode As Long

   Screen.MousePointer = vbHourglass

   For DMJ = DMJE.DDag To DMJE.DJaar
      For Datum = 0 To DatumLijst.ListCount - 1
         For AndereDatum = 0 To DatumLijst.ListCount - 1
            If Not Datum = AndereDatum Then
               Periode = Val(DatumElement(DatumLijst.List(Datum), DMJ))
               AnderePeriode = Val(DatumElement(DatumLijst.List(AndereDatum), DMJ))
               If AnderePeriode > Periode Then Verwissel DatumLijst, Datum, AndereDatum
            End If
         Next AndereDatum
      Next Datum
   Next DMJ
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure maakt de lijst van alle deelname datums.
Private Sub WerkLijstBij()
On Error GoTo Fout
Dim Bestand As String
Dim Datum As String

   Screen.MousePointer = vbHourglass
   
   DatumLijst.Clear
   Bestand = Dir$("*.dat", vbArchive Or vbHidden Or vbSystem)
   Do Until Bestand = vbNullString
      Datum = Left$(Bestand, InStr(Bestand, ".") - 1)
      DatumLijst.AddItem DatumElement(Datum, DMJE.DDag) & "-" & DatumElement(Datum, DMJE.DMaand) & "-" & DatumElement(Datum, DMJE.DJaar)
      Bestand = Dir$()
   Loop
   
   SorteerLijst
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub AnnulerenKnop_Click()
On Error GoTo Fout
   Unload DeelnameDatumsVenster
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
   DeelnameDatumsVenster.Left = (HoofdVenster.Width / 2) - (DeelnameDatumsVenster.Width / 2)
   DeelnameDatumsVenster.Top = (HoofdVenster.Height / 3) - (DeelnameDatumsVenster.Height / 2)
   
   WerkLijstBij
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



'Deze procedure geeft de opdracht om de deelnames voor de geselecteerde datum te tonen.
Private Sub ToonDeelnamesKnop_Click()
On Error GoTo Fout
   DeelnameVenster.VanDatum = DatumLijst.List(DatumLijst.ListIndex)
   DeelnameVenster.TotDatum = DatumLijst.List(DatumLijst.ListIndex)

   Unload DeelnameDatumsVenster

   DeelnameVenster.Show
   DeelnameVenster.SetFocus
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


