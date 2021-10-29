Attribute VB_Name = "Autofilltest1"
Sub AutofillData()

'Calculating Parameters...

Dim NCells As Integer
Dim NPersons As Integer
Dim NYears As Integer
Dim DataInput As Worksheet
Sheets("Data Input").Range("O2") = "=COUNTA(N:N)"
NCells = Worksheets("Data Input").Range("O2").Value
Sheets("Data Input").Range("P2") = "=O2-1"
NPersons = Worksheets("Data Input").Range("P2").Value
Sheets("Data Input").Range("Q2") = "=SUM(IF(FREQUENCY(INDIRECT(""N2:N""&O2),INDIRECT(""N2:N""&O2))>0,1))"
NYears = Worksheets("Data Input").Range("Q2").Value
Set DataInput = Worksheets("Data Input")

Dim MaxYear As Integer, MinYear As Integer, YearPaste As Integer
MaxYear = Application.WorksheetFunction.Max(Worksheets("Data Input").Range("N:N"))
MinYear = Application.WorksheetFunction.Min(Worksheets("Data Input").Range("N:N"))
YearPaste = 2

Do While MinYear <= MaxYear

DataInput.Range("R" & YearPaste) = MinYear

MinYear = MinYear + 1
YearPaste = YearPaste + 1

Loop

'AutofillTime!

If Sheets("CVD Prediction").Range("C3") <> 0 Or Sheets("Non-CVD Prediction").Range("B3") <> 0 Then
    Sheets("CVD Prediction").Range("A3:CF100000").ClearContents
    Sheets("Non-CVD Prediction").Range("A3:CF100000").ClearContents
End If

Range("'Non-CVD Prediction'!A2:'Non-CVD Prediction'!CF2").AutoFill Destination:=Range("'Non-CVD Prediction'!A2:'Non-CVD Prediction'!CF" & NCells)
Range("'CVD Prediction'!A2:'CVD Prediction'!CF2").AutoFill Destination:=Range("'CVD Prediction'!A2:'CVD Prediction'!CF" & NCells)

'Dim sum stuff....yummm

Dim NCopies As Integer
Dim ResultsYear As Integer
Dim NonCVDPrediction As Worksheet
Dim CVDPrediction As Worksheet
Dim ArthritisRisk As Double
Dim COPDRisk As Double
Dim DepressionRisk As Double
Dim DiabetesRisk As Double
Dim CVDRisk As Double
Dim StartRow As Integer
Dim EndRow As Integer
Dim NCount As Integer
Dim ArthritisTotalCost As Double
Dim COPDTotalCost As Double
Dim DepressionTotalCost As Double
Dim DiabetesTotalCost As Double
Dim CVDTotalCost As Double
Dim TC As Double
Dim HC As Double
Dim HDLC As Double
Dim SBP As Double
Dim HT As Double
Dim Smoke As Double
Dim Glucose As Double
Dim Diabetic As Double
Dim BMI As Double
Dim WHR As Double
Dim HealthScore As Double
NCount = NPersons / NYears
Set NonCVDPrediction = Worksheets("Non-CVD Prediction")
Set CVDPrediction = Worksheets("CVD Prediction")


'Calculate Averages

'Year 1

ArthritisRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "U"), NonCVDPrediction.Cells(NCount + 1, "U")))
COPDRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "V"), NonCVDPrediction.Cells(NCount + 1, "V")))
DepressionRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "W"), NonCVDPrediction.Cells(NCount + 1, "W")))
DiabetesRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "X"), NonCVDPrediction.Cells(NCount + 1, "X")))
CVDRisk = Application.Average(CVDPrediction.Range(CVDPrediction.Cells(2, "BP"), CVDPrediction.Cells(NCount + 1, "BP")))
ArthritisTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "BY"), NonCVDPrediction.Cells(NCount + 1, "BY")))
COPDTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "BZ"), NonCVDPrediction.Cells(NCount + 1, "BZ")))
DepressionTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "CA"), NonCVDPrediction.Cells(NCount + 1, "CA")))
DiabetesTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "CB"), NonCVDPrediction.Cells(NCount + 1, "CB")))
CVDTotalCost = Application.Average(CVDPrediction.Range(CVDPrediction.Cells(2, "CF"), CVDPrediction.Cells(NCount + 1, "CF")))
TC = Application.Average(DataInput.Range(DataInput.Cells(2, "D"), DataInput.Cells(NCount + 1, "D")))
HC = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "M"), NonCVDPrediction.Cells(NCount + 1, "M"))) * 100
HDLC = Application.Average(DataInput.Range(DataInput.Cells(2, "E"), DataInput.Cells(NCount + 1, "E")))
SBP = Application.Average(DataInput.Range(DataInput.Cells(2, "F"), DataInput.Cells(NCount + 1, "F")))
HT = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "N"), NonCVDPrediction.Cells(NCount + 1, "N"))) * 100
Smoke = Application.Average(DataInput.Range(DataInput.Cells(2, "H"), DataInput.Cells(NCount + 1, "H"))) * 100
Glucose = Application.Average(DataInput.Range(DataInput.Cells(2, "I"), DataInput.Cells(NCount + 1, "I")))
Diabetic = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(2, "R"), NonCVDPrediction.Cells(NCount + 1, "R"))) * 100
BMI = Application.Average(DataInput.Range(DataInput.Cells(2, "J"), DataInput.Cells(NCount + 1, "J")))
WHR = Application.Average(DataInput.Range(DataInput.Cells(2, "K"), DataInput.Cells(NCount + 1, "K")))
HealthScore = Application.Average(DataInput.Range(DataInput.Cells(2, "M"), DataInput.Cells(NCount + 1, "M")))
Sheets("Results").Range("B2") = ArthritisRisk
Sheets("Results").Range("B3") = COPDRisk
Sheets("Results").Range("B4") = DepressionRisk
Sheets("Results").Range("B5") = DiabetesRisk
Sheets("Results").Range("B6") = CVDRisk
Sheets("Results").Range("B9") = ArthritisTotalCost
Sheets("Results").Range("B10") = COPDTotalCost
Sheets("Results").Range("B11") = DepressionTotalCost
Sheets("Results").Range("B12") = DiabetesTotalCost
Sheets("Results").Range("B13") = CVDTotalCost
Sheets("Results").Range("B16") = TC
Sheets("Results").Range("B17") = HC
Sheets("Results").Range("B18") = HDLC
Sheets("Results").Range("B19") = SBP
Sheets("Results").Range("B20") = HT
Sheets("Results").Range("B21") = Smoke
Sheets("Results").Range("B22") = Glucose
Sheets("Results").Range("B23") = Diabetic
Sheets("Results").Range("B24") = BMI
Sheets("Results").Range("B25") = WHR
Sheets("Results").Range("B26") = HealthScore
ResultsYear = Application.WorksheetFunction.Min(Worksheets("Data Input").Range("N:N"))
Sheets("Results").Range("C1") = ResultsYear
Sheets("Results").Range("C8") = ResultsYear
Sheets("Results").Range("C15") = ResultsYear
NCopies = NYears - 1
Sheets("Results").Range("A1:B26").Copy
ResultsYear = Application.WorksheetFunction.Max(Worksheets("Data Input").Range("N:N"))
StartRow = NCells - NCount
EndRow = NCells

'Other Years

Do While NCopies > 0

Range(Cells(NCopies * 27 + 1, 1), Cells(NCopies * 27 + 1, 2)).PasteSpecial
Sheets("Results").Range("C" & NCopies * 27 + 1) = ResultsYear
Sheets("Results").Range("C" & NCopies * 27 + 8) = ResultsYear
Sheets("Results").Range("C" & NCopies * 27 + 15) = ResultsYear
ArthritisRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "U"), NonCVDPrediction.Cells(EndRow, "U")))
COPDRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "V"), NonCVDPrediction.Cells(EndRow, "V")))
DepressionRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "W"), NonCVDPrediction.Cells(EndRow, "W")))
DiabetesRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "X"), NonCVDPrediction.Cells(EndRow, "X")))
CVDRisk = Application.Average(CVDPrediction.Range(CVDPrediction.Cells(StartRow, "BP"), CVDPrediction.Cells(EndRow, "BP")))
ArthritisTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "BY"), NonCVDPrediction.Cells(EndRow, "BY")))
COPDTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "BZ"), NonCVDPrediction.Cells(EndRow, "BZ")))
DepressionTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "CA"), NonCVDPrediction.Cells(EndRow, "CA")))
DiabetesTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "CB"), NonCVDPrediction.Cells(EndRow, "CB")))
CVDTotalCost = Application.Average(CVDPrediction.Range(CVDPrediction.Cells(StartRow, "CF"), CVDPrediction.Cells(EndRow, "CF")))
TC = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "D"), DataInput.Cells(EndRow, "D")))
HC = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "M"), NonCVDPrediction.Cells(EndRow, "M"))) * 100
HDLC = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "E"), DataInput.Cells(EndRow, "E")))
SBP = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "F"), DataInput.Cells(EndRow, "F")))
HT = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "N"), NonCVDPrediction.Cells(EndRow, "N"))) * 100
Smoke = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "H"), DataInput.Cells(EndRow, "H"))) * 100
Glucose = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "I"), DataInput.Cells(EndRow, "I")))
Diabetic = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "R"), NonCVDPrediction.Cells(EndRow, "R"))) * 100
BMI = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "J"), DataInput.Cells(EndRow, "J")))
WHR = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "K"), DataInput.Cells(EndRow, "K")))
HealthScore = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "M"), DataInput.Cells(EndRow, "M")))
Sheets("Results").Range("B" & NCopies * 27 + 2) = ArthritisRisk
Sheets("Results").Range("B" & NCopies * 27 + 3) = COPDRisk
Sheets("Results").Range("B" & NCopies * 27 + 4) = DepressionRisk
Sheets("Results").Range("B" & NCopies * 27 + 5) = DiabetesRisk
Sheets("Results").Range("B" & NCopies * 27 + 6) = CVDRisk
Sheets("Results").Range("B" & NCopies * 27 + 9) = ArthritisTotalCost
Sheets("Results").Range("B" & NCopies * 27 + 10) = COPDTotalCost
Sheets("Results").Range("B" & NCopies * 27 + 11) = DepressionTotalCost
Sheets("Results").Range("B" & NCopies * 27 + 12) = DiabetesTotalCost
Sheets("Results").Range("B" & NCopies * 27 + 13) = CVDTotalCost
Sheets("Results").Range("B" & NCopies * 27 + 16) = TC
Sheets("Results").Range("B" & NCopies * 27 + 17) = HC
Sheets("Results").Range("B" & NCopies * 27 + 18) = HDLC
Sheets("Results").Range("B" & NCopies * 27 + 19) = SBP
Sheets("Results").Range("B" & NCopies * 27 + 20) = HT
Sheets("Results").Range("B" & NCopies * 27 + 21) = Smoke
Sheets("Results").Range("B" & NCopies * 27 + 22) = Glucose
Sheets("Results").Range("B" & NCopies * 27 + 23) = Diabetic
Sheets("Results").Range("B" & NCopies * 27 + 24) = BMI
Sheets("Results").Range("B" & NCopies * 27 + 25) = WHR
Sheets("Results").Range("B" & NCopies * 27 + 26) = HealthScore
StartRow = StartRow - NCount
EndRow = EndRow - NCount
ResultsYear = ResultsYear - 1

NCopies = NCopies - 1

Loop

Sheets("Data Input").Range("R2", "R" & NYears + 1).Copy
Sheets("Results").Range("E2", "E" & NYears + 1).PasteSpecial
Sheets("Results").Range("G2", "G" & NYears + 1).PasteSpecial
Sheets("Results").Range("I2", "I" & NYears + 1).PasteSpecial
Sheets("Results").Range("K2", "K" & NYears + 1).PasteSpecial
Sheets("Results").Range("M2", "M" & NYears + 1).PasteSpecial
Sheets("Results").Range("O2", "O" & NYears + 1).PasteSpecial
Sheets("Results").Range("Q2", "Q" & NYears + 1).PasteSpecial
Sheets("Results").Range("S2", "S" & NYears + 1).PasteSpecial
Sheets("Results").Range("U2", "U" & NYears + 1).PasteSpecial
Sheets("Results").Range("W2", "W" & NYears + 1).PasteSpecial
Sheets("Results").Range("Y2", "Y" & NYears + 1).PasteSpecial
Sheets("Results").Range("AA2", "AA" & NYears + 1).PasteSpecial
Sheets("Results").Range("AC2", "AC" & NYears + 1).PasteSpecial
Sheets("Results").Range("AE2", "AE" & NYears + 1).PasteSpecial
Sheets("Results").Range("AG2", "AG" & NYears + 1).PasteSpecial
Sheets("Results").Range("AI2", "AI" & NYears + 1).PasteSpecial
Sheets("Results").Range("AK2", "AK" & NYears + 1).PasteSpecial
Sheets("Results").Range("AM2", "AM" & NYears + 1).PasteSpecial
Sheets("Results").Range("AO2", "AO" & NYears + 1).PasteSpecial
Sheets("Results").Range("AQ2", "AQ" & NYears + 1).PasteSpecial
Sheets("Results").Range("AS2", "AS" & NYears + 1).PasteSpecial
Sheets("Results").Range("AU2", "AU" & NYears + 1).PasteSpecial

StartRow = NCells - NCount
EndRow = NCells

'''Metric measures

Dim YearCount As Integer, TotalDiseaseCost As Double

YearCount = NYears

Do While YearCount > 0

ArthritisRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "U"), NonCVDPrediction.Cells(EndRow, "U")))
COPDRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "V"), NonCVDPrediction.Cells(EndRow, "V")))
DepressionRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "W"), NonCVDPrediction.Cells(EndRow, "W")))
DiabetesRisk = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "X"), NonCVDPrediction.Cells(EndRow, "X")))
CVDRisk = Application.Average(CVDPrediction.Range(CVDPrediction.Cells(StartRow, "BP"), CVDPrediction.Cells(EndRow, "BP")))
ArthritisTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "BY"), NonCVDPrediction.Cells(EndRow, "BY")))
COPDTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "BZ"), NonCVDPrediction.Cells(EndRow, "BZ")))
DepressionTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "CA"), NonCVDPrediction.Cells(EndRow, "CA")))
DiabetesTotalCost = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "CB"), NonCVDPrediction.Cells(EndRow, "CB")))
CVDTotalCost = Application.Average(CVDPrediction.Range(CVDPrediction.Cells(StartRow, "CF"), CVDPrediction.Cells(EndRow, "CF")))
TC = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "D"), DataInput.Cells(EndRow, "D")))
HC = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "M"), NonCVDPrediction.Cells(EndRow, "M"))) * 100
HDLC = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "E"), DataInput.Cells(EndRow, "E")))
SBP = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "F"), DataInput.Cells(EndRow, "F")))
HT = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "N"), NonCVDPrediction.Cells(EndRow, "N"))) * 100
Smoke = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "H"), DataInput.Cells(EndRow, "H"))) * 100
Glucose = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "I"), DataInput.Cells(EndRow, "I")))
Diabetic = Application.Average(NonCVDPrediction.Range(NonCVDPrediction.Cells(StartRow, "R"), NonCVDPrediction.Cells(EndRow, "R"))) * 100
BMI = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "J"), DataInput.Cells(EndRow, "J")))
WHR = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "K"), DataInput.Cells(EndRow, "K")))
HealthScore = Application.Average(DataInput.Range(DataInput.Cells(StartRow, "M"), DataInput.Cells(EndRow, "M")))
TotalDiseaseCost = ArthritisTotalCost + COPDTotalCost + DepressionTotalCost + DiabetesTotalCost + CVDTotalCost

Sheets("Results").Range("F" & YearCount + 1) = ArthritisRisk
Sheets("Results").Range("H" & YearCount + 1) = COPDRisk
Sheets("Results").Range("J" & YearCount + 1) = DepressionRisk
Sheets("Results").Range("L" & YearCount + 1) = DiabetesRisk
Sheets("Results").Range("N" & YearCount + 1) = CVDRisk
Sheets("Results").Range("P" & YearCount + 1) = ArthritisTotalCost
Sheets("Results").Range("R" & YearCount + 1) = COPDTotalCost
Sheets("Results").Range("T" & YearCount + 1) = DepressionTotalCost
Sheets("Results").Range("V" & YearCount + 1) = CVDTotalCost
Sheets("Results").Range("X" & YearCount + 1) = TC
Sheets("Results").Range("Z" & YearCount + 1) = HC
Sheets("Results").Range("AB" & YearCount + 1) = HDLC
Sheets("Results").Range("AD" & YearCount + 1) = SBP
Sheets("Results").Range("AF" & YearCount + 1) = HT
Sheets("Results").Range("AH" & YearCount + 1) = Smoke
Sheets("Results").Range("AJ" & YearCount + 1) = Glucose
Sheets("Results").Range("AL" & YearCount + 1) = Diabetic
Sheets("Results").Range("AN" & YearCount + 1) = BMI
Sheets("Results").Range("AP" & YearCount + 1) = WHR
Sheets("Results").Range("AR" & YearCount + 1) = HealthScore
Sheets("Results").Range("AT" & YearCount + 1) = TotalDiseaseCost


Sheets("Results").Range("F1") = "Arthritis Risk"
Sheets("Results").Range("H1") = "COPD Risk"
Sheets("Results").Range("J1") = "Depression Risk"
Sheets("Results").Range("L1") = "Diabetes Risk"
Sheets("Results").Range("N1") = "CVD Risk"
Sheets("Results").Range("P1") = "Arthritis Cost"
Sheets("Results").Range("R1") = "COPD Cost"
Sheets("Results").Range("T1") = "Depression Cost"
Sheets("Results").Range("V1") = "CVD Cost"
Sheets("Results").Range("X1") = "Total Cholesterol"
Sheets("Results").Range("Z1") = "% High Cholesterol"
Sheets("Results").Range("AB1") = "HDL-C"
Sheets("Results").Range("AD1") = "SBP"
Sheets("Results").Range("AF1") = "% High Blood Pressure"
Sheets("Results").Range("AH1") = "% Smoke"
Sheets("Results").Range("AJ1") = "Glucose"
Sheets("Results").Range("AL1") = "% (Pre-)Diabetic"
Sheets("Results").Range("AN1") = "BMI"
Sheets("Results").Range("AP1") = "WHR"
Sheets("Results").Range("AR1") = "Health Score"
Sheets("Results").Range("AT1") = "Total Disease Cost"

StartRow = StartRow - NCount
EndRow = EndRow - NCount

YearCount = YearCount - 1

Loop

Sheets("Results").Range("E1:AS9").Font.Color = RGB(255, 255, 255)

Sheets("Results").Range("E1:AS9").Select

With Selection
    .ShrinkToFit = True
End With

ActiveWindow.ScrollColumn = 1

MsgBox ("Complete!")

End Sub

