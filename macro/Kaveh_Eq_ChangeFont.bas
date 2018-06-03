Option Explicit

'**************************************************************************
'Macro Name: Kaveh_Eq_ChangeFont()
'Version: 1.0
'
'Description: This Macro change the font of all equations in a MS Word document.
'
'This Macro is developed by Kaveh Bakhtiyari, and it is copyrighted.
'You can use this macro for free, but do not remove this copyright header.
'www.bakhtiyari.com
'**************************************************************************

Sub Kaveh_Eq_ChangeFont()
	If Not ActiveDocument.Saved Then
		If vbYes = MsgBox("Do you want to save your document?", vbYesNo, "Save document") Then
			ActiveDocument.Save
		End If
	End If

	Dim FontName As String
	FontName = InputBox("Enter your font name (e.g. Latin Modern Math)", "Math Font Name", "Latin Modern Math")

	Dim i As Integer
	For i = 1 To ActiveDocument.OMaths.Count
		ActiveDocument.OMaths.Item(i).Range.Select
		'Selection.Font.Size = 12
		Selection.Font.Name = FontName
	Next i

	MsgBox i & " equations were updated."
End Sub
