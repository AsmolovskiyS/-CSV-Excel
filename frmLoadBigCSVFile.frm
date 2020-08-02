VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoadBigCSVFile 
   Caption         =   "Загрузщик CSV файлоф с большим количеством строк"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9705.001
   OleObjectBlob   =   "frmLoadBigCSVFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoadBigCSVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBRazdelitel_Change()
    CBVibFil.SetFocus
End Sub

'Изменение имени сохраняемого файла
Private Sub CBSave_Click()
    Dim TempFileName As Variant
    
    TempFileName = GetNameFile(LoadBigCSVFile.SaveFile, "FoldName")
    TempFileName = Application.GetSaveAsFilename(TempFileName, "Двоичный файл Excel (*.xlsb),*.xlsb,Файл Excel (*.xlsx),*.xlsx,Файл Excel с поддержкой макросов (*.xlsm),*.xlsm", 0, "Укажите имя и место сохранения файла")
    If TempFileName <> False Then
        LoadBigCSVFile.SaveFile = TempFileName
        LSaveFile = TempFileName
    End If
End Sub

Private Sub CBVibFil_Click()
    Dim TempFileName As Variant
    TempFileName = Application.GetOpenFilename(, , "Выберите файл", , False)
    If TempFileName <> False Then
        LoadBigCSVFile.Razdelitel = TypeRazdelintelja()
        LoadBigCSVFile.FileName = TempFileName
        LFileName.Caption = TempFileName
        LoadBigCSVFile.SaveFile = LoadBigCSVFile.GetNameFile(TempFileName, "FoldName") & ".xlsb"
        LSaveFile.Caption = LoadBigCSVFile.SaveFile
        Call LoadBigCSVFile.PodschetKolRows
        LKolRow.Caption = CStr(LoadBigCSVFile.KolRow)
        LKolCol.Caption = CStr(LoadBigCSVFile.KolCol)
        TBFirstLine.Value = LoadBigCSVFile.FirstLine
        Call EnableElements
        If LoadBigCSVFile.KolRow < 1000000 Then
            TBKolRowOnList.Value = CStr(LoadBigCSVFile.KolRow)
        End If
        Call RaschetListov
        If Len(LoadBigCSVFile.FirstLine) > 100 Then
            Frame1.Height = 94
            TBFirstLine.Height = 32
        End If
    End If
End Sub

Function TypeRazdelintelja() As String
    Select Case CBRazdelitel.ListIndex
        Case 0
            TypeRazdelintelja = ";"
        Case 1
            TypeRazdelintelja = vbTab
        Case 2
            TypeRazdelintelja = "."
        Case 3
            TypeRazdelintelja = ","
    End Select
End Function

Private Sub CBLoad_Click()
    If ProverkaGotovnostiKZagruzke Then
        LoadBigCSVFile.KolListov = CInt(LKolList.Caption)
        LoadBigCSVFile.KolRowOnList = CLng(TBKolRowOnList.Value)
        Call DisableElements
        Call LoadBigCSVFile.StartLoadData
    End If
End Sub

Function ProverkaGotovnostiKZagruzke() As Boolean
    Dim TempKolRow As Long
    If LoadBigCSVFile.FileName <> "False" Then
        If Not IsNumeric(TBKolRowOnList.Value) Then
            MsgBox ("Значение количества строк указано не верно!")
            ProverkaGotovnostiKZagruzke = False
            Exit Function
        Else
            TempKolRow = CLng(TBKolRowOnList.Value)
            TBKolRowOnList.Value = CStr(TempKolRow)
            If TempKolRow < 1 Or TempKolRow > 1048576 Then
                MsgBox ("Значение количсетва строк на лист должно быть в диапозоне от 1 до 1048576!")
                ProverkaGotovnostiKZagruzke = False
                Exit Function
            End If
        End If
     Else
        MsgBox ("Файл не выбран, выберите файл!")
        ProverkaGotovnostiKZagruzke = False
        Exit Function
     End If
     ProverkaGotovnostiKZagruzke = True
End Function

Private Sub TBKolRowOnList_Change()
    Call RaschetListov
End Sub

Sub RaschetListov()
    Dim TempKolRow As Long
    
    If Len(TBKolRowOnList.Value) > 0 Then
        TempKolRow = CLng(TBKolRowOnList.Value)
        TBKolRowOnList.Value = CStr(TempKolRow)
        If TempKolRow > 1 Then
            If TempKolRow > LoadBigCSVFile.KolRow Then
                TBKolRowOnList.Value = CStr(LoadBigCSVFile.KolRow)
            End If
            LKolList.Caption = CStr(WorksheetFunction.RoundUp((LoadBigCSVFile.KolRow - 1) / (TempKolRow - 1), 0))
            Exit Sub
        Else
            TBKolRowOnList.Value = ""
        End If
    End If
    LKolList.Caption = " "
End Sub

Private Sub TBKolRowOnList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii.Value = 0
    End If
End Sub


Private Sub UserForm_Initialize()
    CBRazdelitel.AddItem "Точка с запятой"
    CBRazdelitel.AddItem "Табуляция"
    CBRazdelitel.AddItem "Точка"
    CBRazdelitel.AddItem "Запятая"
    CBRazdelitel.ListIndex = 0
    CBRazdelitel.Style = fmStyleDropDownList
    
    LoadBigCSVFile.FileName = "False"
    
    LoadBigCSVFile.KolRowOnList = 1000000
    TBKolRowOnList.Value = CStr(LoadBigCSVFile.KolRowOnList)
    
    Call DisableElements
    
    ProgressBar.Width = 0.1
End Sub

Sub EnableElements()
    CBLoad.Enabled = True
    TBKolRowOnList.Enabled = True
    CBSave.Enabled = True
End Sub

Sub DisableElements()
    CBLoad.Enabled = False
    TBKolRowOnList.Enabled = False
    CBSave.Enabled = False
End Sub


