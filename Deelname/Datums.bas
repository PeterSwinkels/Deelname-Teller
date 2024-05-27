Attribute VB_Name = "DatumsModule"
'Deze module bevat de procedures voor het werken met datums.
Option Explicit

'Deze opsomming definieert de elementen waaruit een datum is opgebouwd.
Public Enum DMJE
   DDag      'Dag element.
   DMaand    'Maand element.
   DJaar     'Jaar element.
End Enum

'Deze structuur definieert de elementen van een datum.
Public Type DatumStr
   Dag As String     'Definieert het dag element.
   Maand As String   'Definieert het maand element.
   Jaar As String    'Definieert het jaar element.
End Type
'Deze procedure stuurt het aantal dagen in de opgegeven maand terug.
Private Function AantalDagenInMaand(Maand As Long, Jaar As Long) As Long
On Error GoTo Fout
Dim AantalDagen As Long

   AantalDagen = 0

   Select Case Maand
      Case 1, 3, 5, 7, 8, 10, 12
         AantalDagen = 31
      Case 2
         If IsSchrikkeljaar(Jaar) Then AantalDagen = 28
      Case 4, 6, 9, 11
         AantalDagen = 30
   End Select

EindeProcedure:
   AantalDagenInMaand = AantalDagen
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function



'Deze procedure controleert of de opgegeven datum binnnen de opgegeven periode valt en stuurt het resultaat terug.
Public Function BinnenPeriode(Datum As String, VanDatum As String, TotDatum As String) As Boolean
On Error GoTo Fout
Dim Dag As Long
Dim Jaar As Long
Dim Maand As Long
Dim NaDanVan As Boolean
Dim TotDag As Long
Dim TotJaar As Long
Dim TotMaand As Long
Dim VanDag As Long
Dim VanJaar As Long
Dim VanMaand As Long
Dim VoorDanTot As Boolean

   Dag = Val(DatumElement(Datum, DMJE.DDag))
   Jaar = Val(DatumElement(Datum, DMJE.DJaar))
   Maand = Val(DatumElement(Datum, DMJE.DMaand))
   NaDanVan = False
   TotDag = Val(DatumElement(TotDatum, DMJE.DDag))
   TotJaar = Val(DatumElement(TotDatum, DMJE.DJaar))
   TotMaand = Val(DatumElement(TotDatum, DMJE.DMaand))
   VanDag = Val(DatumElement(VanDatum, DMJE.DDag))
   VanJaar = Val(DatumElement(VanDatum, DMJE.DJaar))
   VanMaand = Val(DatumElement(VanDatum, DMJE.DMaand))
   VoorDanTot = False

   If Jaar > VanJaar Then
      NaDanVan = True
   ElseIf Jaar = VanJaar Then
      If Maand > VanMaand Then
         NaDanVan = True
      ElseIf Maand = VanMaand Then
         If Dag >= VanDag Then NaDanVan = True
      End If
   End If

   If Jaar < TotJaar Then
      VoorDanTot = True
   ElseIf Jaar = TotJaar Then
      If Maand < TotMaand Then
         VoorDanTot = True
      ElseIf Maand = TotMaand Then
         If Dag <= TotDag Then VoorDanTot = True
      End If
   End If

EindeProcedure:
   BinnenPeriode = NaDanVan And VoorDanTot
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function
'Deze procedure stuurt het opgegeven element van de opgegeven datum terug.
Public Function DatumElement(OpgegevenDatum As String, DMJ As DMJE) As String
On Error GoTo Fout
Dim Dag As String
Dim Datum As String
Dim Element As String
Dim Jaar As String
Dim Maand As String
Dim Positie As Long

      
   Datum = Trim$(OpgegevenDatum)
   Element = vbNullString

   Positie = InStr(Datum, "-")
   If Positie > 0 Then
      Dag = Left$(Datum, Positie - 1)
      Datum = Mid$(Datum, Positie + 1)
      
      Positie = InStr(Datum, "-")
      Maand = Left$(Datum, Positie - 1)
      Datum = Mid$(Datum, Positie + 1)
      
      Jaar = Datum
   ElseIf Len(Datum) = 8 Then
      Dag = Left$(Datum, 2)
      Maand = Mid$(Datum, 3, 2)
      Jaar = Mid$(Datum, 5, 4)
   End If
 
   If Len(Jaar) < 3 Then Jaar = Mid$(CStr(Year(Date)), 1, Len(CStr(Year(Date))) - 2) & Format$(Jaar, "00")
 
   Select Case DMJ
      Case DMJE.DDag
         Element = Dag
      Case DMJE.DMaand
         Element = Maand
      Case DMJE.DJaar
         Element = Jaar
   End Select

EindeProcedure:
   DatumElement = Element
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure stuurt de opgegeven datum in DDMMJJJJ formaat terug.
Public Function DDMMJJJJ(Datum As String) As String
On Error GoTo Fout
Dim DDMMJJJJV As String

   DDMMJJJJV = Format$(DatumElement(Datum, DMJE.DDag), "00") & Format$(DatumElement(Datum, DMJE.DMaand), "00") & Format$(DatumElement(Datum, DMJE.DJaar), "0000")

EindeProcedure:
   DDMMJJJJ = DDMMJJJJV
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure controleert of de opgegeven datum geldig is.
Public Function IsGeldigeDatum(Datum As String) As Boolean
On Error GoTo Fout
Dim Dag As Long
Dim IsGeldig As Boolean
Dim Jaar As Long
Dim Maand As Long

   IsGeldig = False

   If Not Datum = vbNullString Then
      Dag = CLng(Val(DatumElement(Datum, DMJE.DDag)))
      Maand = CLng(Val(DatumElement(Datum, DMJE.DMaand)))
      Jaar = CLng(Val(DatumElement(Datum, DMJE.DJaar)))
            
      If Maand >= 1 And Maand <= 12 Then
         If Dag >= 1 And Dag <= AantalDagenInMaand(Maand, Jaar) Then
            If Jaar >= 1900 And Jaar <= 2099 Then
               IsGeldig = True
            End If
         End If
      End If
   End If
   
EindeProcedure:
   IsGeldigeDatum = IsGeldig
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function




'Deze procedure controleert of het opgegeven jaar een schrikkeljaar is en stuurt het resultaat terug.
Private Function IsSchrikkeljaar(Jaar As Long) As Boolean
On Error GoTo Fout
Dim IsSchrikkel As Boolean

   IsSchrikkel = False
   
   If Jaar Mod 4 = 0 Then
      If Jaar Mod 100 = 0 Then
         If Jaar Mod 400 = 0 Then IsSchrikkel = True
      Else
         IsSchrikkel = True
      End If
   End If

EindeProcedure:
   IsSchrikkeljaar = IsSchrikkel
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure vraagt om een begindatum en einddatum.
Public Sub VraagPeriode(ByRef VanDatum As String, ByRef TotDatum As String)
On Error GoTo Fout
   VanDatum = InputBox$("Van: (dd-mm-jjjj):", , "01-01-1900")
   TotDatum = InputBox$("Tot: (dd-mm-jjjj):", , "31-12-2099")
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


