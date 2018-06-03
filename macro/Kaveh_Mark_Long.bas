Option Explicit

Sub Kaveh_Mark_Long()
    
    '**************************************************************************
    'Macro Name: Kaveh_Mark_Long()
    'Version: 1.3
    '
    'Update 1.3:
    'Update notification has been added.
    'Document Save question box is added.
    '
    'Update 1.2:
    'Ignore field data (e.g. EndNote References)
    'It can be switched off/on by altering blnRemoveField
    'Enter character is removed to count the number of words correctly.
    '
    'Update 1.1:
    'Track Changes issue is fixed.
    '
    '
    'Description: This Macro counts and marks the sentences with more than a specific
    'number of words (e.g. 30). It helps authors to recognize the long sentences thus
    'to simplify them.
    '
    '
    'This Macro is developed by Kaveh Bakhtiyari, and it is copyrighted.
    'You can use this macro for free, but do not remove this copyright header.
    'www.bakhtiyari.com
    '**************************************************************************
    
    Dim CurrentVersion As Single
    CurrentVersion = 1.3
    
    On Error Resume Next
    
    Dim LatestVersion As Single
    Dim DownloadURL As String
    
    Dim iMyCount As Integer
    Dim iWords As Integer
    Dim strWords As String
    Dim iCurrent As Integer
    Dim iMaxSent As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim blnRemoveField As Boolean
    
    Dim MySent As Range
    Dim rgMaxSent As Range
    Dim objFlD As Field
    Dim i As Integer
    Dim blnTrackChanges As Boolean
    
    blnRemoveField = True
    
    Dim doc As MSXML2.DOMDocument60
    Set doc = New MSXML2.DOMDocument60
    doc.async = False
    If doc.Load("http://www.bakhtiyari.com/version.xml") Then
        LatestVersion = CSng(doc.SelectSingleNode("/AppData/KavehMarkLong/Version").Text)
        DownloadURL = doc.SelectSingleNode("/AppData/KavehMarkLong/url").Text
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

    blnTrackChanges = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False
    
    'Reset counter
    iMyCount = 0
    iMaxSent = 0
    iCurrent = 0
    
    'Set number of words
    strWords = InputBox("Enter minimum number of words per sentence (e.g. 30)", "Minimum Mumber of Words", "30")
    If IsNumeric(strWords) Then
        iWords = CInt(strWords)
    Else
        iWords = 30
    End If
    
    'Remove the punctuations from the sentence
    Dim sSent As String
    Dim StrReplace(25) As String
    StrReplace(0) = ","
    StrReplace(1) = "."
    StrReplace(2) = ";"
    StrReplace(3) = "'"
    StrReplace(4) = """"
    StrReplace(5) = "!"
    StrReplace(6) = "~"
    StrReplace(7) = "?"
    StrReplace(8) = "!"
    StrReplace(9) = ":"
    StrReplace(10) = "!"
    StrReplace(11) = "{"
    StrReplace(12) = "}"
    StrReplace(13) = "["
    StrReplace(14) = "]"
    StrReplace(15) = "<"
    StrReplace(16) = ">"
    StrReplace(17) = "("
    StrReplace(18) = ")"
    StrReplace(19) = "|"
    StrReplace(20) = "\"
    StrReplace(21) = "/"
    StrReplace(22) = Chr(13)
    StrReplace(23) = ""
    StrReplace(24) = ""
    
    For Each MySent In ActiveDocument.Sentences
        If blnRemoveField = True Then
           For Each objFlD In MySent.Fields
                objFlD.ShowCodes = True
            Next
        End If
               
        sSent = MySent.Text
        
        If blnRemoveField = True Then
            For Each objFlD In MySent.Fields
                sSent = Replace(sSent, objFlD.Code.Text, "")
                objFlD.ShowCodes = False
            Next
        End If
        
        For i = 0 To UBound(StrReplace)
            sSent = Replace(sSent, StrReplace(i), "")
        Next
        sSent = Replace(sSent, "  ", " ")
        sSent = Trim(sSent)
        
        'If sCheck.Words.Count > iWords Then
        iCurrent = UBound(Split(sSent, " ")) + 1
        'iCurrent = MySent.ComputeStatistics(wdStatisticWords)
        
        If iCurrent > iWords Then
            If iCurrent > iMaxSent And iCurrent < 500 Then
                iMaxSent = iCurrent
                Set rgMaxSent = MySent
            End If
            
            MySent.Font.Color = wdColorRed
            iMyCount = iMyCount + 1
        End If
                        
    Next
    If iMyCount > 0 Then
        rgMaxSent.Font.Color = wdColorGreen
        ActiveDocument.Comments.Add rgMaxSent, "[MACRO] The longest sentence in your document with " & iMaxSent & " words."
    End If
    
    ActiveDocument.TrackRevisions = blnTrackChanges
    ActiveWindow.View.ShowFieldCodes = False
    
    MsgBox iMyCount & " sentences longer than " & iWords & " (" & Int(((iMyCount * 100) / ActiveDocument.Sentences.Count)) & _
           "%) words out of " & ActiveDocument.Sentences.Count & " sentenses. Maximum: " & iMaxSent & " words."

    If vbYes = MsgBox("Would you like to check our website?", vbYesNo, "Do you like it?") Then
        ActiveDocument.FollowHyperlink Address:="http://www.bakhtiyari.com"
    End If

End Sub
