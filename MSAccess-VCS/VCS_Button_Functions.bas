Attribute VB_Name = "VCS_Button_Functions"
Option Compare Database

' function to call update functions from button clicks
Function subUpdateBtn(btnFunction As String)
    ' do this every time
    ' loadVCS

    Select Case btnFunction
        Case "importSourceBtn" ', "exportFormsBtn", "resetFormsBtn"
            Debug.Print "button worked: " & btnFunction
            ImportAllSource(True) ' will skip importing tables
        Case "exportSourceBtn"
            Debug.Print "button worked2: " & btnFunction
            ExportAllSource(True) ' will skip exporting tables
		Case "ImportProjectBtn"
			Debug.Print "button worked3: " & btnFunction
            ImportProject(True) ' will skip importing tables
		Case Else
			MsgBox "Current function doesn't yet exist: " & btnFunction
    End Select

End Function
