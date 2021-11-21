Option Explicit

Sub PP_Footer_Outline()

    '**************************************************************************
    'Macro Name: PP_Footer_Outline()
    'Version: 1.0
    '
    'Description: This macro writes the names of the sections in the Powerpoint footer,
	'and it changes the font/color of the current section in order to show which section
	'the slide belongs to.
	'In addition, it writes the active slide number and total number of slides dynamically
	'(e.g. x of X) without hardcoding the total number.
    '
    '
    'This Macro is developed by Kaveh Bakhtiyari, and it is copyrighted.
    'You can use this macro for free, but do not remove this copyright header.
    'www.bakhtiyari.com
    '**************************************************************************

Dim osld As Slide
Dim oshp As Shape
Dim str_Subsection As String
Dim str_SectionList As String
Dim i As Long

Dim b_found As Boolean

If ActivePresentation.SectionProperties.Count > 0 Then

For Each osld In ActivePresentation.Slides
    osld.HeadersFooters.Footer.Visible = True
    osld.HeadersFooters.SlideNumber.Visible = msoTrue

    If osld.CustomLayout.Name = "Title Slide" Then
            If osld.Shapes.HasTitle Then str_Subsection = osld.Shapes.Title.TextFrame.TextRange
    End If

    For Each oshp In osld.Shapes
    
'*********************************************************
'Add slide numbers and adding the total number of slides
'*********************************************************
        If Left(oshp.Name, 12) = "Slide Number" Then
            oshp.TextFrame.TextRange.Text = osld.SlideNumber & " of " & ActivePresentation.Slides.Count
        End If
        
'*********************************************************
'Add powerpoint navigation outline based on section names
'*********************************************************
        If oshp.Type = msoPlaceholder Then
            If oshp.PlaceholderFormat.Type = ppPlaceholderFooter Then
                                
                str_SectionList = ""
                For i = 1 To ActivePresentation.SectionProperties.Count
                    If osld.sectionIndex = i Then
                        lngStart = Len(str_SectionList)
                        lngEnd = Len(ActivePresentation.SectionProperties.Name(i)) + 1
                    End If
                    str_SectionList = str_SectionList & ActivePresentation.SectionProperties.Name(i)
                    If i < ActivePresentation.SectionProperties.Count Then
                        str_SectionList = str_SectionList & " - "
                    End If
                Next i

                With oshp.TextFrame.TextRange
                    .Font.Bold = msoFalse
                    .Font.Italic = msoFalse
                    .Font.Color.RGB = RGB(200, 200, 200)
                    .Font.Size = 18
                End With
                
                oshp.TextFrame.TextRange = str_SectionList
                
                With oshp.TextFrame.TextRange.Characters(1, lngStart - 1)
                    .Font.Bold = msoFalse
                    .Font.Italic = msoFalse
                    .Font.Color.RGB = RGB(200, 200, 200)
                End With
                With oshp.TextFrame.TextRange.Characters(lngStart, lngEnd)
                    .Font.Bold = msoTrue
                    .Font.Italic = msoTrue
                    .Font.Color.RGB = RGB(255, 0, 0)
                End With
                With oshp.TextFrame.TextRange.Characters(lngStart + lngEnd + 1)
                    .Font.Bold = msoFalse
                    .Font.Italic = msoFalse
                    .Font.Color.RGB = RGB(200, 200, 200)
                End With
                
            End If
        End If
    Next oshp
Next osld
End If
End Sub
