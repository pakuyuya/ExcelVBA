'
' ディレクトリを選ぶ
'
Private Function ChooseDir(Optional defaultPath = vbNullString) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ""
    
        If .Show = True Then
            ChooseDir = .SelectedItems(1)
        End If
        ChooseDir = defaultPath
    End With
End Function