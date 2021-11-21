Option Explicit

Sub ENChangeCitationColor()
    
    '**************************************************************************
    'Macro Name: ENChangeCitationColor()
    'Version: 1.1
    '
    'Update 1.1:
    'Update Notification has been added.
    'Document Save question box is added.
    '
    '
    'Description: This Macro marks (DarkBlue) the inline citations of EndNote.
	'EndNote software does not have any built-in feature to change the color of in-line citations in a Word document.
	'This macro changes the color of all citations made by EndNote, and makes the references in a document clear.
    '
    'This Macro is developed by Kaveh Bakhtiyari, and it is copyrighted.
    'You can use this macro for free, but do not remove this copyright heading.
    'www.bakhtiyari.com
    '**************************************************************************
    
    Dim CurrentVersion As Single
    CurrentVersion = 1.1
    
    On Error Resume Next
    
    Dim LatestVersion As Single
    Dim DownloadURL As String

    Dim doc As MSXML2.DOMDocument60
    Set doc = New MSXML2.DOMDocument60
    doc.async = False
    If doc.Load("http://www.bakhtiyari.com/version.xml") Then
        LatestVersion = CSng(doc.SelectSingleNode("/AppData/ENChangeCitationColor/Version").Text)
        DownloadURL = doc.SelectSingleNode("/AppData/ENChangeCitationColor/url").Text
        If LatestVersion > CurrentVersion Then
            If vbYes = MsgBox("Currently, you are running version " & CurrentVersion & "." & Chr(13) & "There is a new version " & _
                LatestVersion & " available at " & DownloadURL & "." & Chr(13) & "Would you like to update?", vbYesNo, "Update available") Then
                ActiveDocument.FollowHyperlink Address:=DownloadURL
            End If
        End If
    End If
    
    If Not ActiveDocument.Saved Then
        If vbYes = MsgBox("Do you want to save your document?", vbYesNo, "Save document") Then
           ActiveDocument.Save
        End If
    End If

    Dim blnTrackChanges As Boolean
    blnTrackChanges = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False

    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Color = Word.WdColor.wdColorDarkBlue
    End With
    With Selection.Find
        .Text = "^19 ADDIN EN.CITE"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    
    ActiveDocument.TrackRevisions = blnTrackChanges

    If vbYes = MsgBox("Would you like to check our website?", vbYesNo, "Do you like it?") Then
        ActiveDocument.FollowHyperlink Address:="http://bakhtiyari.com"
    End If

End Sub
