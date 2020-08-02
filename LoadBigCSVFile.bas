Attribute VB_Name = "LoadBigCSVFile"
Option Explicit

Public Razdelitel As String
Public FileName As String
Public KolRow, KolCol, KolRowOnList As Long
Public KolListov
Private strBuffer As String
Public FirstLine As String
Public SaveFile As String
Public Book As Workbook

Public Sub Start()
    frmLoadBigCSVFile.Show (0)
End Sub

'������������ ���������� ����� � ��������
Public Sub PodschetKolRows()
    Dim lfAnsi As String
    Dim F As Integer
    Dim buffer() As Byte
    Dim bufPos As Long
    Dim lineCount As Long
    Dim copyFirstLine As Boolean
    
    lfAnsi = StrConv(vbLf, vbFromUnicode)
    F = FreeFile ' ����� ��������� ����� ����
    Open FileName For Binary Access Read As #F ' ��������� ���� ��� ������
    ReDim buffer(LOF(F)) ' ������� ������ �� ������� �����
    Get #F, , buffer ' ��������� ���� � ������
    strBuffer = buffer '�������� � ������
    Erase buffer '������� ������
   
    bufPos = 1
    FirstLine = False
    Do Until bufPos = 0
        bufPos = InStrB(bufPos, strBuffer, lfAnsi)
        If bufPos > 0 Then
            If Not copyFirstLine Then
                FirstLine = StrConv(LeftB(strBuffer, bufPos - 2), vbUnicode)
                KolCol = UBound(Split(FirstLine, Razdelitel)) + 1
                copyFirstLine = True
            End If
            lineCount = lineCount + 1
            bufPos = bufPos + 1
        End If
    Loop
    KolRow = lineCount + 1
    Close #F
End Sub

Sub StartLoadData()
    'Call UskorenieVkl ' ��������� ���������� ������ �������� �������� ����� �� ����
    Dim ArrRows() As String
    Dim ArrRange() As Variant
    Dim i, n As Long
    Dim List As Integer
    Dim TempRowsList As Long
    Dim RowsOnPx As Long
    Dim OldPos, Pos As Variant
    Dim FirstLine As Boolean
    Dim T0 As Single
    
    T0 = Timer()
    RowsOnPx = WorksheetFunction.RoundUp(KolRow / (frmLoadBigCSVFile.ProgressBarRamka.Width - 1), 0)
    
    Call SendLog("���������� �����")
    Call AddNewBook
    
    Call SendLog("��������� ������")
    strBuffer = StrConv(strBuffer, vbUnicode)
    
    Call SendLog("�������� ������")
    List = 1
    TempRowsList = TekusheeKolRowsOnList(List)
    Call ArrayReDim(ArrRange, TempRowsList)
    n = -1
    
    OldPos = 0
    Pos = 1
    FirstLine = False
    i = 0
    Do Until Pos = 0
        i = i + 1
        '������� ��� � ���������� ������ �� �����
        If i Mod RowsOnPx = 0 Or i = KolRow Or n + 1 = TempRowsList - 1 Or i = 1 Then
            frmLoadBigCSVFile.LTekList.Caption = CStr(List)
            frmLoadBigCSVFile.LZagRow.Caption = CStr(i)
            frmLoadBigCSVFile.ProgressBar.Width = Round((frmLoadBigCSVFile.ProgressBarRamka.Width - 1) / KolRow * i, 0)
            DoEvents
        End If
        
        Pos = InStr(Pos, strBuffer, vbLf)
        
        If Not FirstLine Then
            OldPos = Pos
            Pos = Pos + 1
            FirstLine = True
        Else
            n = n + 1
            If Pos > 0 Then
                Call ZapolnenieMassiva(ArrRange, Mid(strBuffer, OldPos + 1, Pos - OldPos - 2), n)
                OldPos = Pos
                Pos = Pos + 1
            End If
            If Pos = 0 Then
                Call ZapolnenieMassiva(ArrRange, Mid(strBuffer, OldPos + 1, Len(strBuffer) - OldPos - 1), n)
            End If
        End If
        
        If n = TempRowsList - 1 Then
            Call SendLog("������� ������ �� ����")
            Book.Worksheets(CStr(List)).Select
            Range(Cells(2, 1), Cells(TempRowsList + 1, KolCol)).Value = ArrRange
            n = -1
            If List < KolListov Then
                List = List + 1
                TempRowsList = TekusheeKolRowsOnList(List)
                Call ArrayReDim(ArrRange, TempRowsList)
                Call SendLog("�������� ������")
            End If
        End If
        
        If i = KolRow Then
            Exit Do
        End If
    Loop
      
    Call SendLog("���������� �����")
    Book.Save
    
    'Call UskorenieVikl
    Call SendLog("���������� �����")
    Call SendLog("�������� ������ ��������� �� " & Format$(Timer() - T0, "0.0#") & " c.")
    MsgBox ("�������� ������ ���������")
End Sub

Private Sub ZapolnenieMassiva(ByRef Arr() As Variant, ByVal Txt As String, ByVal Position As Long)
    Dim ArrCols() As String
    Dim i As Long
    
    ArrCols = Split(Txt, Razdelitel)
    For i = LBound(ArrCols) To UBound(ArrCols)
        Arr(Position, i) = ArrCols(i)
    Next i
End Sub

Private Function TekusheeKolRowsOnList(ByVal List As Integer) As Long
    If List < KolListov Or KolListov = 1 Then
        TekusheeKolRowsOnList = KolRowOnList - 1 '����� 1, ��� ��� ������ ������ ��� �����
    Else
        TekusheeKolRowsOnList = KolRow - KolRowOnList - (List - 2) * (KolRowOnList - 1) '����� 1, ��� ��� ������ ������ ��� �����
    End If
End Function

Private Sub SendLog(ByVal Text As String)
    frmLoadBigCSVFile.LLog.Caption = Text
    DoEvents
End Sub

Private Sub ArrayReDim(ByRef Arr() As Variant, ByVal RowOnList As Long)
    Erase Arr
    ReDim Arr(RowOnList - 1, KolCol - 1) '����� 1 ������ ��� ���� � ����
End Sub

Private Sub AddNewBook()
    Dim TempFormat As String
    Dim TempList As Worksheet
    Dim i As Integer
    Dim TempCols() As String
    Dim Cols() As String
     
    '������� � ��������� �����
    Set Book = Workbooks.Add
    TempFormat = GetNameFile(SaveFile, "Extn")
    Select Case TempFormat
        Case "xlsb"
            LoadBigCSVFile.Book.SaveAs FileName:=SaveFile, FileFormat:=50
        Case "xlsx"
            LoadBigCSVFile.Book.SaveAs FileName:=SaveFile, FileFormat:=51
        Case "xlsm"
            LoadBigCSVFile.Book.SaveAs FileName:=SaveFile, FileFormat:=52
    End Select
    
    '�������������� ����� ��������
    ReDim Cols(0, KolCol - 1)
    TempCols = Split(FirstLine, Razdelitel)
    For i = LBound(TempCols) To UBound(TempCols)
        Cols(0, i) = TempCols(i)
    Next i
    
    '������� ������ ���������� ������ � ������� �����������
    '������� � �������� �������, ��� �� ������ ���� ������ �� ����
    For i = KolListov To 1 Step -1
        Set TempList = Book.Worksheets.Add
        TempList.Name = CStr(i)
        '��������� ����� ��������
        TempList.Range(Cells(1, 1), Cells(1, KolCol)).Value = Cols
    Next i
    Application.DisplayAlerts = False
    Book.Worksheets("����1").Delete
    Book.Worksheets("����2").Delete
    Book.Worksheets("����3").Delete
    Application.DisplayAlerts = True
End Sub

'������� ��� ��������� ������ �� ������� ����� ����� (FN)
'(������ ��� FN = "C:\Users\Asmolovskiy\Docunents\Test.txt")
'Part:  "Fold" - ����� ���� � ����� (������ "C:\Users\Asmolovskiy\Docunents\t")
'       "FoldName" - ���� � ��� ����� ��� ���������� (������ "C:\Users\Asmolovskiy\Docunents\Test")
'       "Name" - ��� ����� ��� ����������� (������ "Test")
'       "NameExtn" - ��� ����� � ����������� (������ "Test.txt")
'       "Extn" - ��������� ����� (������ "txt")
Function GetNameFile(ByVal FN As String, ByVal Part As String) As String
    Select Case Part
        Case "Fold"
            GetNameFile = Left(FN, InStrRev(FN, "\"))
        Case "FoldName"
            GetNameFile = Left(FN, InStrRev(FN, ".") - 1)
        Case "Name"
            GetNameFile = Mid(FN, InStrRev(FN, "\") + 1, Len(FN) - InStrRev(FN, "."))
        Case "NameExtn"
            GetNameFile = Right(FN, Len(FN) - InStrRev(FN, "\"))
        Case "Extn"
            GetNameFile = Right(FN, Len(FN) - InStrRev(FN, "."))
        Case Else
            GetNameFile = ""
    End Select
End Function

'�������� ���������
Sub UskorenieVkl()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
End Sub

'��������� ���������
Sub UskorenieVikl()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub


