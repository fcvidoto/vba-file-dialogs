VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Incluir a Referencia do Microsoft Office Object

'Seleciona a pasta e ve os arquivos
Function SelectFolderFile() As String
    
    Dim fDialog As Office.FileDialog
    Dim arquivo As Scripting.File
    Dim fso As New Scripting.FileSystemObject
    Set fDialog = Access.Application.FileDialog(2)
    
    'Exibe a janela
    With fDialog
        .AllowMultiSelect = False
        .Title = "Salve o arquivo como"
        .InitialFileName = "relatorio.xlsx"
        '.Filters.Add "XLSM", "*.xlsm", 1
        
        If .Show = True Then
            'Set arquivo = Nothing
            'Set arquivo = fso.GetFile(.SelectedItems(1))
            SelectFolderFile = fDialog.SelectedItems(1)
        Else
            SelectFolderFile = ""
        End If
    End With
    
    '---------------------------------------
    'Se não for selecionado um endereço do report
    If SelectFolderFile = "" Then Exit Function
    
    '---------------------------------------
    'coloca a extensao ".xlsx"
    If Right(SelectFolderFile, 5) <> ".xlsx" Then
      SelectFolderFile = SelectFolderFile & ".xlsx"
    End If
    
End Function

'Seleciona uma pasta
Public Function SelectFolder() As String

    Dim fd As Object
    
    Set fd = Access.Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Selecione a Pasta"
        If .Show = -1 Then       'User pressed action button
            DoEvents
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""
        End If
    End With
    
    Set fd = Nothing
End Function

'File Dialog que seleciona um arquivo XLSXM (Excel com macros) para uma string
Function SelectMacroFile() As String
    
    Dim fDialog As Office.FileDialog
    Dim arquivo As Scripting.File
    Dim fso As New Scripting.FileSystemObject
    Set fDialog = Access.Application.FileDialog(msoFileDialogFilePicker)
    
    'Exibe a janela
    With fDialog
        .AllowMultiSelect = False
        .Title = "Selecione o Arquivo"
        .Filters.Add "XLSM", "*.xlsm", 1
        
        If .Show = True Then
            Set arquivo = Nothing
            Set arquivo = fso.GetFile(.SelectedItems(1))
            SelectMacroFile = arquivo
        Else
            SelectMacroFile = ""
        End If
    End With

End Function

'File Dialog que seleciona um arquivo XLSX (Excel) para uma string
Function SelectFile(ByVal titulo As String) As String
    
    Dim fDialog As Office.FileDialog
    Dim arquivo As Scripting.File
    Dim fso As New Scripting.FileSystemObject
    Set fDialog = Access.Application.FileDialog(msoFileDialogFilePicker)
    
    'Incluí cálculos automáticos a planilha
    'Application.Calculation = xlCalculationAutomatic
    
    'Exibe a janela
    With fDialog
        .AllowMultiSelect = False
        .Title = titulo
        .Filters.Add "XLS, XLSX, XLSM, XLSB", "*.xlsx;*.xls;*.xlsm;*.xlsb", 1
        
        If .Show = True Then
            Set arquivo = Nothing
            Set arquivo = fso.GetFile(.SelectedItems(1))
            SelectFile = arquivo
        Else
            SelectFile = ""
        End If
    End With

End Function


