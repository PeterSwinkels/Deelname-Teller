VERSION 5.00
Begin VB.Form DeelnameVenster 
   Caption         =   "Deelnames"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5895
   Icon            =   "deelname.frx":0000
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   14.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   49.125
   Begin VB.PictureBox KnoppenBalk 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1560
      ScaleHeight     =   3.063
      ScaleMode       =   4  'Character
      ScaleWidth      =   35.125
      TabIndex        =   4
      Top             =   2520
      Width           =   4215
      Begin VB.PictureBox SorteerKnoppenBalk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         ScaleHeight     =   3.063
         ScaleMode       =   4  'Character
         ScaleWidth      =   21.125
         TabIndex        =   8
         Top             =   0
         Width           =   2535
         Begin VB.CommandButton SorteerOpDeelname 
            Caption         =   "&Deelname"
            Height          =   495
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton SorteerOpLidNummer 
            Caption         =   "&Lidnummer"
            Height          =   495
            Left            =   0
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label SorteerDeelnamesOpLabel 
            Caption         =   "Sorteer deelnames op:"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.CommandButton SluitenKnop 
         Cancel          =   -1  'True
         Caption         =   "&Sluiten"
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox DeelnameLijst 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label PeriodeLabel 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   405
   End
   Begin VB.Label TotaalDeelnamesVeld 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label TotaalLabel 
      Caption         =   "Totaal:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   615
   End
   Begin VB.Menu PrintenMenu 
      Caption         =   "&Printen"
   End
   Begin VB.Menu ExporterenMenu 
      Caption         =   "&Exporteren"
   End
   Begin VB.Menu VerversLijstMenu 
      Caption         =   "&Ververs Lijst"
   End
   Begin VB.Menu InformatieMenu 
      Caption         =   "&Informatie"
   End
End
Attribute VB_Name = "DeelnameVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat de lijst waarin het totale aantal deelnames voor ieder lidnummer wordt getoond.
Option Explicit

'Deze opsomming definieert een lijst van methodes om de deelnamelijst te sorteren.
Private Enum SorteermethodesE
   SAantalDeelnames   'Sorteert op het aantal deelnames.
   SLidnummer         'Sorteert op lidnummer.
End Enum

Private Sorteermethode As SorteermethodesE   'Bevat de sorteermethode voor de deelnamelijst.
Private TotDatumV As String                  'Bevat de einddatum van de door de gebruiker opgegeven periode.
Private VanDatumV As String                  'Bevat de begindatum van de door de gebruiker opgegeven periode.
'Deze procedure stuurt het item op de opgegeven rij en kolom terug.
Private Function LijstItem(ItemRij As Long, ItemKolom As Long) As String
On Error GoTo Fout
Dim Item As String
Dim Kolom As Long

   Item = DeelnameLijst.List(ItemRij)
   
   Kolom = 0
   Do Until Kolom = ItemKolom
      Item = Trim$(Mid$(Item, InStr(Item, "  ")))
      Kolom = Kolom + 1
   Loop
   If InStr(Item, "  ") > 0 Then Item = Left$(Item, InStr(Item, "  "))
 
EindeProcedure:
   LijstItem = Item
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure stelt de nieuwe begin datum in.
Public Property Let VanDatum(NieuweVanDatum As String)
On Error GoTo Fout
   VanDatumV = NieuweVanDatum
EindeProcedure:
   Exit Property

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Property

'Deze procedure stelt de nieuwe einddatum in.
Public Property Let TotDatum(NieuweTotDatum As String)
On Error GoTo Fout
   TotDatumV = NieuweTotDatum
EindeProcedure:
   Exit Property

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Property

'Deze procedure sorteert de lid nummers op het aantal deelnames.
Private Sub SorteerLijstOpDeelnames()
On Error GoTo Fout
Dim AndereRij As Long
Dim Rij As Long

   Screen.MousePointer = vbHourglass

   For Rij = 1 To DeelnameLijst.ListCount - 1
      For AndereRij = 1 To DeelnameLijst.ListCount - 1
         If Val(LijstItem(AndereRij, ItemKolom:=2)) > Val(LijstItem(Rij, ItemKolom:=2)) Then Verwissel DeelnameLijst, Rij, AndereRij
      Next AndereRij
   Next Rij

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure werkt de getoonde gegevens in dit venster bij.
Private Sub WerkGegevensBij()
On Error GoTo Fout
Dim Lidnummer As Long
Dim TotaalDeelnames As Long

   If VanDatumV = TotDatumV Then
      PeriodeLabel = "Datum: " & VanDatumV
      SorteerKnoppenBalk.Visible = False
   Else
      PeriodeLabel = "Periode: " & VanDatumV & " - " & TotDatumV
      SorteerKnoppenBalk.Visible = True
   End If
 
   BerekenDeelnames VanDatumV, TotDatumV

   Screen.MousePointer = vbHourglass

   TotaalDeelnames = 0
   DeelnameLijst.Clear
   DeelnameLijst.AddItem Space$(5) & "Lidnummer" & Space$(24) & "Deelname"
   For Lidnummer = LBound(Deelname()) To UBound(Deelname())
      If Deelname(Lidnummer) > 0 Then
         DeelnameLijst.AddItem Space$(5) & Format$(Lidnummer + 1, "0000") & Space$(30) & Deelname(Lidnummer)
         TotaalDeelnames = TotaalDeelnames + Val(Deelname(Lidnummer))
      End If
   Next Lidnummer
   
   TotaalDeelnamesVeld = TotaalDeelnames
   
   If Sorteermethode = SorteermethodesE.SAantalDeelnames Then SorteerLijstOpDeelnames

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure exporteert deelnamelijst naar het tekst bestand Deelname.txt.
Private Sub ExporterenMenu_Click()
On Error GoTo Fout
Dim BestandH As Integer
Dim Rij As Long

   Screen.MousePointer = vbHourglass
   BestandH = FreeFile()
   Open "Deelname.txt" For Output Lock Read Write As BestandH
      Print #BestandH, PeriodeLabel
      Print #BestandH, DeelnameLijst.List(0)
      Print #BestandH, String$(80, "-")
      For Rij = 1 To DeelnameLijst.ListCount
         Print #BestandH, DeelnameLijst.List(Rij)
      Next Rij
      Print #BestandH, String$(80, "-")
      Print #BestandH, "Totaal: "; TotaalDeelnamesVeld
   Close 1

   MsgBox "De lijst is geëxporteerd naar Deelname.txt", vbInformation

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht de getoonde gegevens bij te werken.
Private Sub Form_Activate()
On Error GoTo Fout
   WerkGegevensBij
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft opdracht om de programma informatie te tonen.
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
Private Sub Form_Load()
On Error GoTo Fout
   DeelnameVenster.Left = MenuVenster.Left + MenuVenster.Width + 256
   DeelnameVenster.Top = MenuVenster.Top
   DeelnameVenster.Width = HoofdVenster.Width - DeelnameVenster.Left - 512
   DeelnameVenster.Height = (HoofdVenster.Height - DeelnameVenster.Top - 1024)
   
   Sorteermethode = SorteermethodesE.SLidnummer
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure past dit venster aan de nieuwe grootte aan.
Private Sub Form_Resize()
On Error Resume Next
   DeelnameLijst.Width = DeelnameVenster.ScaleWidth - 2
   DeelnameLijst.Height = DeelnameVenster.ScaleHeight - KnoppenBalk.Height - 2
   
   KnoppenBalk.Left = DeelnameVenster.ScaleWidth - KnoppenBalk.Width - 2
   KnoppenBalk.Top = DeelnameVenster.ScaleHeight - KnoppenBalk.Height - 1
   
   TotaalDeelnamesVeld.Left = TotaalLabel.Left + TotaalLabel.Width
   TotaalDeelnamesVeld.Top = DeelnameLijst.Top + DeelnameLijst.Height
   TotaalLabel.Left = 10
   TotaalLabel.Top = DeelnameLijst.Top + DeelnameLijst.Height
End Sub

'Deze procedure stelt de printer in en drukt de deelnamelijst af.
Private Sub PrintenMenu_Click()
On Error GoTo Fout
Dim Kolom As Long
Dim PrinterProbleem As Boolean
Dim Rij As Long

   StelPrinterIn PrinterProbleem
   
   If Not PrinterProbleem Then
      Screen.MousePointer = vbHourglass
   
      Printer.CurrentY = 2
      Printer.CurrentX = 3: Printer.Print PeriodeLabel
      For Rij = 0 To DeelnameLijst.ListCount - 1
         For Kolom = 1 To 2
            Printer.CurrentX = Choose(Kolom, 3, 43)
            Printer.Print LijstItem(Rij, Kolom);
         Next Kolom
         Printer.Print
         If Rij = 1 Then Printer.Print String$(80, "-")
      Next Rij
      Printer.Print String$(80, "-")
      Printer.CurrentX = 3: Printer.Print "Totaal: "; TotaalDeelnamesVeld
      Printer.EndDoc
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub





'Deze procedure sluit dit venster.
Private Sub SluitenKnop_Click()
On Error GoTo Fout
   Unload DeelnameVenster
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure geeft de opdracht de lijst gesorteerd op het aantal deelnames te tonen.
Private Sub SorteerOpDeelname_Click()
On Error GoTo Fout
   Sorteermethode = SorteermethodesE.SAantalDeelnames
   WerkGegevensBij
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure geeft de opdracht de lijst gesorteerd op het lidnummer te tonen.
Private Sub SorteerOpLidNummer_Click()
On Error GoTo Fout
   Sorteermethode = SorteermethodesE.SLidnummer
   WerkGegevensBij
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure geeft opdracht om de deelnamelijst bij te werken.
Private Sub VerversLijstMenu_Click()
On Error GoTo Fout
   WerkGegevensBij
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


