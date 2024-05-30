Attribute VB_Name = "KernModule"
'Deze module bevat alle kernprocedures en variabelen voor dit programma.
Option Explicit

Private Deelgenomen(0 To 9999) As Boolean   'Bevat de deelnamelijst voor een datum waarin aangegeven wordt of een lid deel genomen heeft.
Public Deelname(0 To 9999) As Long         'Bevat het aantal deelnames binnen een bepaalde periode.

'Deze procedure berekent het totale aantal deelnames voor alle lid nummers binnen de opgegeven periode.
Public Sub BerekenDeelnames(VanDatum As String, TotDatum As String)
On Error GoTo Fout
Dim Bestand As String
Dim Datum As String
Dim Lidnummer As Long

   If Not (VanDatum = vbNullString Or TotDatum = vbNullString) Then
      Erase Deelname

      Screen.MousePointer = vbHourglass

      Bestand = Dir$("*.dat", vbArchive Or vbHidden Or vbSystem)
      Do Until Bestand = vbNullString Or DoEvents() = 0
         Datum = Left$(Bestand, InStr(Bestand, ".") - 1)
         
         If BinnenPeriode(Datum, VanDatum, TotDatum) Then
            LaadDeelnames Datum
            For Lidnummer = LBound(Deelgenomen()) To UBound(Deelgenomen())
               If Deelgenomen(Lidnummer) Then Deelname(Lidnummer) = Deelname(Lidnummer) + 1
            Next Lidnummer
         End If
         
         Bestand = Dir$()
      Loop
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure bewaart de deelnames onder opgegeven datum.
Private Sub BewaarDeelnames(Datum As String)
On Error GoTo Fout
Dim BestandH As Integer
Dim Buffer As String
Dim Lidnummer As Long

   Screen.MousePointer = vbHourglass

   Buffer = vbNullString
   For Lidnummer = LBound(Deelgenomen()) To UBound(Deelgenomen())
      Buffer = Buffer & Chr$(Abs(Deelgenomen(Lidnummer)))
   Next Lidnummer

   BestandH = FreeFile()
   Open DDMMJJJJ(Datum) & ".dat" For Output Lock Read Write As BestandH
      Print #BestandH, Buffer;
   Close BestandH

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure handelt eventuele fouten af.
Public Function HandelFoutAf(Optional VraagVorigeKeuzeOp As Boolean = False) As Long
Dim Bericht As String
Dim Foutcode As Long
Static Keuze As Long

   Screen.MousePointer = vbDefault

   Bericht = Err.Description
   Foutcode = Err.Number
   Err.Clear
      
   If Not VraagVorigeKeuzeOp Then
      Bericht = Bericht & vbCr & "Foutcode: " & CStr(Foutcode)
   
      Keuze = MsgBox(Bericht, vbAbortRetryIgnore Or vbExclamation)
   End If
   
   HandelFoutAf = Keuze
   
   If Keuze = vbAbort Then End
End Function

'Deze procedure controleert of het opgegeven lidnummer geldig is en stuurt het resultaat terug.
Private Function IsGeldigLidnummer(Lidnummer As String) As Boolean
On Error GoTo Fout
Dim IsGeldig As Boolean

   IsGeldig = False

   If CStr(CLng(Val(Lidnummer))) = Lidnummer Then
      If Val(Lidnummer) < 1 Or Val(Lidnummer) > 10000 Then
         MsgBox "Het lidnummer moet tussen de 1 en 10000 zijn.", vbExclamation
      ElseIf Val(Lidnummer) = 0 Then
         MsgBox "Het lidnummer kan geen nul zijn.", vbExclamation
      Else
         IsGeldig = True
      End If
   End If
   
EindeProcedure:
   IsGeldigLidnummer = IsGeldig
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure laadt de deelnames voor de opgegeven datum.
Private Sub LaadDeelnames(Datum As String)
On Error GoTo Fout
Dim BestandH As Integer
Dim Buffer As String
Dim Lidnummer As Long

   Erase Deelgenomen

   Screen.MousePointer = vbHourglass

   BestandH = FreeFile()
   Open DDMMJJJJ(Datum) & ".dat" For Binary Lock Read Write As BestandH
      If LOF(BestandH) = 0 Then
         Close BestandH
         Kill DDMMJJJJ(Datum) & ".dat"
      Else
         Buffer = Input$(Abs(UBound(Deelgenomen()) - LBound(Deelgenomen())) + 1, BestandH)
        
         For Lidnummer = LBound(Deelgenomen()) To UBound(Deelgenomen())
            Deelgenomen(Lidnummer) = CBool(Asc(Mid$(Buffer, Lidnummer + 1, 1)))
         Next Lidnummer
      End If
   Close BestandH

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure wordt uitgevoerd wanneer het programma wordt gestart.
Public Sub Main()
On Error GoTo Fout

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   Erase Deelgenomen, Deelname
    
   HoofdVenster.Show
   Do While DoEvents() > 1
   Loop
   
   Unload HoofdVenster
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure stelt de printer in.
Public Sub StelPrinterIn(PrinterProbleem As Boolean)
On Error GoTo PrinterFout
Dim Bericht As String
Dim Foutcode As Long
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   
   PrinterProbleem = False
   
   Printer.ColorMode = vbPRCMMonochrome
   Printer.Copies = 1
   Printer.Duplex = vbPRDPSimplex
   Printer.Font.Bold = True
   Printer.Font.Italic = False
   Printer.Font.Name = "Lucida Console"
   Printer.Font.Size = 10
   Printer.Font.Underline = False
   Printer.Font.Weight = 400
   Printer.Orientation = vbPRORPortrait
   Printer.PaperSize = vbPRPSA4
   Printer.PrintQuality = vbPRPQMedium
   Printer.RightToLeft = False
   Printer.ScaleMode = vbCharacters
   Printer.TrackDefault = True
   Printer.Zoom = 1
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

PrinterFout:
   Screen.MousePointer = vbDefault
   Bericht = Err.Description
   Foutcode = Err.Number
   PrinterProbleem = True

   Bericht = Bericht & vbCr & "Foutcode: " & Foutcode
   Keuze = MsgBox(Bericht, vbAbortRetryIgnore Or vbExclamation)
   Select Case Keuze
      Case vbAbort
         Unload HoofdVenster
      Case vbIgnore
         Resume EindeProcedure
      Case vbRetry
         Screen.MousePointer = vbHourglass
         Resume
   End Select
End Sub

'Deze procedure toont het totale aantal deelnames voor het opgegeven lidnummer binnen een bepaalde periode.
Public Sub ToonLidnummerDeelname()
On Error GoTo Fout
Dim IngevoerdLidnummer As String
Dim Lidnummer As Long
Dim TotDatum As String
Dim VanDatum As String

   Do While DoEvents() > 0
      IngevoerdLidnummer = InputBox$("Lidnummer:")
      If IngevoerdLidnummer = vbNullString Then
         Exit Do
      ElseIf IsGeldigLidnummer(IngevoerdLidnummer) Then
         Exit Do
      End If
   Loop
   
   If Not IngevoerdLidnummer = vbNullString Then
      Lidnummer = Val(IngevoerdLidnummer)
   
      VraagPeriode VanDatum, TotDatum
      If Not (VanDatum = vbNullString Or TotDatum = vbNullString) Then
         BerekenDeelnames VanDatum, TotDatum
         MsgBox "Lidnummer " & Lidnummer & " heeft " & Deelname(Lidnummer - 1) & " keer deelgenomen.", vbInformation
      End If
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure toont de informatie over dit programma.
Public Sub ToonProgrammainformatie()
On Error GoTo Fout
   MsgBox App.Comments, vbInformation, App.Title & " v" & App.Major & "." & App.Minor & App.Revision & ", door: " & App.CompanyName
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure vraagt de gebruiker om deelnames te verwijderen.
Public Sub VerwijderDeelnames()
On Error GoTo Fout
Dim Datum As String
Dim IngevoerdLidnummer As String
Dim Lidnummer As Long

   Datum = InputBox$("Datum (dd-mm-jjjj):", App.Title & " - Verwijderen", Day(Date) & "-" & Month(Date) & "-" & Year(Date))
   If IsGeldigeDatum(Datum) Then
      LaadDeelnames Datum

      Do
         Do
            IngevoerdLidnummer = InputBox$("Lidnummer:", App.Title & " - Verwijderen")
            If IngevoerdLidnummer = vbNullString Then Exit Do
         Loop Until IsGeldigLidnummer(IngevoerdLidnummer)
        
         If Not IngevoerdLidnummer = vbNullString Then
            Lidnummer = Val(IngevoerdLidnummer)
             
            If Deelgenomen(Lidnummer - 1) Then
               Deelgenomen(Lidnummer - 1) = False
            Else
               MsgBox "Het opgegeven lid heeft niet deelgenomen op de opgegeven datum.", vbExclamation
            End If
         End If
      Loop Until IngevoerdLidnummer = vbNullString Or DoEvents() = 0
      
      BewaarDeelnames Datum
   ElseIf Not Datum = vbNullString Then
      MsgBox "Ongeldige datum.", vbExclamation
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



'Deze procedure verwijdert de deelname gegevens voor opgegeven periode.
Public Sub VerwijderDeelnamesVoorPeriode()
On Error GoTo Fout
Dim Bestand As String
Dim Datum As String
Dim Keuze As Long
Dim TotDatum As String
Dim VanDatum As String

   VraagPeriode VanDatum, TotDatum
   If Not (VanDatum = vbNullString Or TotDatum = vbNullString) Then
      Keuze = MsgBox("Deelnamelijsten verwijderen?", vbYesNo Or vbQuestion)
      If Keuze = vbYes Then
         Erase Deelname
         Screen.MousePointer = vbHourglass
      
         Bestand = Dir$("*.dat", vbArchive Or vbHidden Or vbSystem)
         Do Until Bestand = vbNullString
            Datum = Left$(Bestand, InStr(Bestand, ".") - 1)
            If BinnenPeriode(Datum, VanDatum, TotDatum) Then Kill Bestand
            Bestand = Dir$()
         Loop
      End If
   End If
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure verwisselt de gegevens op de twee opgegeven rijen in de opgegeven lijst met elkaar.
Public Sub Verwissel(Lijst As ListBox, Index1 As Long, Index2 As Long)
On Error GoTo Fout
Dim Rij1 As String
Dim Rij2 As String

   Rij1 = Lijst.List(Index1)
   Rij2 = Lijst.List(Index2)
   
   Lijst.List(Index1) = Rij2
   Lijst.List(Index2) = Rij1
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure vraagt de gebruiker om deelnames in te voeren.
Public Sub VoerDeelnamesIn()
On Error GoTo Fout
Dim Datum As String
Dim IngevoerdLidnummer As String
Dim Lidnummer As Long

   Datum = InputBox$("Datum (dd-mm-jjjj):", App.Title & " - Invoeren", Day(Date) & "-" & Month(Date) & "-" & Year(Date))
   If IsGeldigeDatum(Datum) Then
      LaadDeelnames Datum

      Do
         Do
            IngevoerdLidnummer = InputBox$("Lidnummer:", App.Title & " - Invoeren")
            If IngevoerdLidnummer = vbNullString Then Exit Do
         Loop Until IsGeldigLidnummer(IngevoerdLidnummer)
        
         If Not IngevoerdLidnummer = vbNullString Then
            Lidnummer = Val(IngevoerdLidnummer)
             
            If Deelgenomen(Lidnummer - 1) Then
               MsgBox "Het opgegeven lid heeft al deelgenomen op de opgegeven datum.", vbExclamation
            Else
               Deelgenomen(Lidnummer - 1) = True
            End If
         End If
      Loop Until IngevoerdLidnummer = vbNullString Or DoEvents() = 0
      
      BewaarDeelnames Datum
   ElseIf Not Datum = vbNullString Then
      MsgBox "Ongeldige datum.", vbExclamation
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



