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

'seleciona a pasta e ve os arquivos - Access
public function selectFolderFile() As String
  Dim fDialog As Object 'Office.FileDialog
  Dim arquivo As Object 'Scripting.File
  Dim fso As Object 'New Scripting.FileSystemObject
  Set fDialog = Access.Application.FileDialog(2)
  With fDialog 'Exibe a janela
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
  If SelectFolderFile = "" Then Exit function 'Se não for selecionado um endereço do report
  '---------------------------------------
  'coloca a extensao ".xlsx"
  If Right(SelectFolderFile, 5) <> ".xlsx" Then SelectFolderFile = SelectFolderFile & ".xlsx"
End Function

'seleciona uma pasta - Access
Public Function selectFolder() As String
  Dim fd As Object
  Set fd = Access.Application.FileDialog(4)
  With fd
    .Title = "Selecione a Pasta"
    If .Show = -1 Then       'User pressed action button
      DoEvents
      selectFolder = .SelectedItems(1)
    Else
      selectFolder = ""
    End If
  End With
  Set fd = Nothing
End Function

'seleciona um arquivo XLSXM (Excel com macros) para uma string - Access
public function selectMacroFile() As String
  Dim fDialog As Object
  Dim fso As Object 'New Scripting.FileSystemObject
  Set fDialog = Access.Application.FileDialog(1) 'msoFileDialogFilePicker
  With fDialog 'Exibe a janela
	  .AllowMultiSelect = False
    .Title = "Selecione o Arquivo"
    .Filters.Add "XLSM", "*.xlsm", 1
    '.Filters.Add "XLS, XLSX, XLSM, XLSB", "*.xlsx;*.xls;*.xlsm;*.xlsb", 1
    If .Show = True Then
      SelectMacroFile = .SelectedItems(1)
    Else
      SelectMacroFile = ""
    End If
  End With
End Function



