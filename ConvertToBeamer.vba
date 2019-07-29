' this code extracts text from PPT(X) and saves to Lyx for latex beamer body
' Provided for free with no guarantees or promises
'WARNING: this will overwrite files in the powerpoint file's folder if there are name collisiona
' Original version for LaTeX beamer by Louis from StackExchange (https://tex.stackexchange.com/users/6321/louis) available here: (https://tex.stackexchange.com/questions/66007/any-way-of-converting-ppt-or-odf-to-beamer-or-org)
' Modified version for LaTeX beamer by Jason Kerwin (www.jasonkerwin.com) on 20 February 2018:
    ' Takes out extra text that printed in the title line
    ' Switches titles to \frametitle{} instead of including them on the \begin{frame} line (sometimes helps with compiling)
    ' Changes the image names to remove original filename, which might have spaces
    ' Removes "\subsection{}" which was printing before each slide
' Adapted for generation of LYX file for LaTeX beamer with metropolis layout by Julian Schuette, 2019
    ' You may want to have Fira fonts installed on your system or choose an alternative font.
'NB you must convert your slides to .ppt format before running this code

Public Const vbDoubleQuote As String = """"

Public Sub ConvertToBeamer()
    Dim objPresentation As Presentation
    Set objPresentation = Application.ActivePresentation

    Dim objSlide As Slide
    Dim objshape As Shape
    Dim objShape4Note As Shape
    Dim hght As Long, wdth As Long
    Dim objGrpItem As Shape

    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2 'Specify stream type - we want To save text/string data.
    fsT.Charset = "utf-8" 'Specify charset For the source text data.
    fsT.Open 'Open the stream And write binary data To the object
    ' fsT.WriteText "special characters: äöüß"

    Dim Name As String, Pth As String, Dest As String, IName As String, ln As String, ttl As String, BaseName As String
    Dim txt As String
    Dim p As Integer, l As Integer, ctr As Integer, i As Integer, j As Integer
    Dim il As Long, cl As Long
    Dim Pgh As TextRange

    Name = Application.ActivePresentation.Name
    p = InStr(Name, ".ppt")
    l = Len(Name)
    If p + 3 = l Then
      Mid(Name, p) = ".lyx"
    Else
      Name = Name & ".lyx"
    End If
    BaseName = Left(Name, l - 4)
    Pth = Application.ActivePresentation.Path
    Dest = Pth & "\" & Name
    ctr = 0
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")

    ' Set objTextFile = objFileSystem.CreateTextFile(Dest, True, True)
    fsT.WriteText "#LyX 2.2 created this file. For more info see http://www.lyx.org/" & vbLf
    fsT.WriteText "\lyxformat 508" & vbLf
    fsT.WriteText "\begin_document" & vbLf
    fsT.WriteText "\begin_header" & vbLf
    fsT.WriteText "\save_transient_properties true" & vbLf
    fsT.WriteText "\origin unavailable" & vbLf
    fsT.WriteText "\textclass beamer" & vbLf
    fsT.WriteText "\begin_preamble" & vbLf
    fsT.WriteText "\usetheme{metropolis}" & vbLf
    fsT.WriteText "\usepackage{textcomp}" & vbLf
    fsT.WriteText "\usepackage{listings}" & vbLf
    fsT.WriteText "\usepackage{fontspec}" & vbLf
    fsT.WriteText "\setmonofont{Fira Code}" & vbLf
    fsT.WriteText "\lstset{basicstyle=\ttfamily}" & vbLf
    fsT.WriteText "\end_preamble" & vbLf
    fsT.WriteText "\use_default_options true" & vbLf
    fsT.WriteText "\maintain_unincluded_children false" & vbLf
    fsT.WriteText "\language english" & vbLf
    fsT.WriteText "\language_package default" & vbLf
    fsT.WriteText "\inputencoding auto" & vbLf
    fsT.WriteText "\fontencoding global" & vbLf
    fsT.WriteText "\font_roman ""default"" ""default""" & vbLf
    fsT.WriteText "\font_sans ""default"" ""default""" & vbLf
    fsT.WriteText "\font_typewriter ""default"" ""default""" & vbLf
    fsT.WriteText "\font_math ""auto"" ""auto""" & vbLf
    fsT.WriteText "\font_default_family default" & vbLf
    fsT.WriteText "\options aspectratio=169,ignorenonframetext" & vbLf
    fsT.WriteText "\use_default_options false" & vbLf
    fsT.WriteText "\maintain_unincluded_children false" & vbLf
    fsT.WriteText "\language english" & vbLf
    fsT.WriteText "\language_package default" & vbLf
    fsT.WriteText "\inputencoding auto" & vbLf
    fsT.WriteText "\fontencoding global" & vbLf
    fsT.WriteText "\font_default_family default" & vbLf
    fsT.WriteText "\use_non_tex_fonts true" & vbLf
    fsT.WriteText "\font_sc false" & vbLf
    fsT.WriteText "\font_osf false" & vbLf
    fsT.WriteText "\font_sf_scale 100 100" & vbLf
    fsT.WriteText "\font_tt_scale 100 100" & vbLf
    fsT.WriteText "\graphics default" & vbLf
    fsT.WriteText "\default_output_format default" & vbLf
    fsT.WriteText "\output_sync 0" & vbLf
    fsT.WriteText "\bibtex_command default" & vbLf
    fsT.WriteText "\index_command default" & vbLf
    fsT.WriteText "\paperfontsize default" & vbLf
    fsT.WriteText "\spacing single" & vbLf
    fsT.WriteText "\use_hyperref false" & vbLf
    fsT.WriteText "\papersize default" & vbLf
    fsT.WriteText "\use_geometry true" & vbLf
    fsT.WriteText "\use_package amsmath 1" & vbLf
    fsT.WriteText "\use_package amssymb 1" & vbLf
    fsT.WriteText "\use_package cancel 1" & vbLf
    fsT.WriteText "\use_package esint 1" & vbLf
    fsT.WriteText "\use_package mathdots 1" & vbLf
    fsT.WriteText "\use_package mathtools 1" & vbLf
    fsT.WriteText "\use_package mhchem 1" & vbLf
    fsT.WriteText "\use_package stackrel 1" & vbLf
    fsT.WriteText "\use_package stmaryrd 1" & vbLf
    fsT.WriteText "\use_package undertilde 1" & vbLf
    fsT.WriteText "\cite_engine basic" & vbLf
    fsT.WriteText "\cite_engine_type default" & vbLf
    fsT.WriteText "\biblio_style plain" & vbLf
    fsT.WriteText "\use_bibtopic false" & vbLf
    fsT.WriteText "\use_indices false" & vbLf
    fsT.WriteText "\paperorientation portrait" & vbLf
    fsT.WriteText "\suppress_date false" & vbLf
    fsT.WriteText "\justification true" & vbLf
    fsT.WriteText "\use_refstyle 1" & vbLf
    fsT.WriteText "\index Index" & vbLf
    fsT.WriteText "\shortcut idx" & vbLf
    fsT.WriteText "\color #008000" & vbLf
    fsT.WriteText "\end_index" & vbLf
    fsT.WriteText "\secnumdepth 3" & vbLf
    fsT.WriteText "\tocdepth 3" & vbLf
    fsT.WriteText "\paragraph_separation indent" & vbLf
    fsT.WriteText "\paragraph_indentation default" & vbLf
    fsT.WriteText "\quotes_language english" & vbLf
    fsT.WriteText "\papercolumns 1" & vbLf
    fsT.WriteText "\papersides 1" & vbLf
    fsT.WriteText "\paperpagestyle default" & vbLf
    fsT.WriteText "\tracking_changes false" & vbLf
    fsT.WriteText "\output_changes false" & vbLf
    fsT.WriteText "\html_math_output 0" & vbLf
    fsT.WriteText "\html_css_as_file 0" & vbLf
    fsT.WriteText "\html_be_strict false" & vbLf
    fsT.WriteText "\end_header" & vbLf
    fsT.WriteText "" & vbLf
    fsT.WriteText "\begin_body" & vbLf
    ' fsT.WriteText "\section{" & Name & "}" & vbLf
    With Application.ActivePresentation.PageSetup
        wdth = .SlideWidth
        hght = .SlideHeight
    End With


    For Each objSlide In objPresentation.Slides
        fsT.WriteText "" & vbLf
        ttl = "No Title"
        If objSlide.Shapes.HasTitle Then
          ttl = objSlide.Shapes.Title.TextFrame.TextRange.Text
        End If
        ' fsT.WriteText "\subsection{" & ttl & "}" & vbLf
        ' fsT.WriteText "\begin{frame}" & vbLf
        ' fsT.WriteText "\frametitle{" & ttl & "}" & vbLf
        fsT.WriteText "\begin_layout FragileFrame" & vbLf
        fsT.WriteText "\begin_inset Argument 4" & vbLf
        fsT.WriteText "status open" & vbLf
        fsT.WriteText "" & vbLf
        fsT.WriteText "\begin_layout Plain Layout" & vbLf
        fsT.WriteText ttl & vbLf
        fsT.WriteText "\end_layout" & vbLf
        fsT.WriteText "" & vbLf
        fsT.WriteText "\end_inset" & vbLf
        fsT.WriteText "" & vbLf
        fsT.WriteText "" & vbLf
        fsT.WriteText "\end_layout" & vbLf
        fsT.WriteText "" & vbLf
        fsT.WriteText "" & vbLf
        fsT.WriteText "\begin_deeper" & vbLf  ' Indent all frame content

        Layout = objSlide.Layout
       
        For Each objshape In objSlide.Shapes

            If objshape.HasTextFrame = True Then
                If Not objshape.TextFrame.TextRange Is Nothing Then
                    
                    il = RealIndent(objshape.TextFrame.TextRange.Paragraphs.IndentLevel)
                    
                    For Each Pgh In objshape.TextFrame.TextRange.Paragraphs

                        If Not objshape.TextFrame.TextRange.Text = ttl Then
                            cl = RealIndent(Pgh.Paragraphs.IndentLevel)
                            txt = Pgh.TrimText
                                                            
                            If Len(txt) > 0 Then
                                If cl > il Then
                                    fsT.WriteText "\begin_deeper" & vbLf
                                    il = cl
                                ElseIf cl < il Then
                                    fsT.WriteText "\end_deeper" & vbLf
                                    il = cl
                                End If
                                
                                If il = 0 Then
                                    fsT.WriteText "\begin_layout Standard" & vbLf
                                    fsT.WriteText txt & vbLf
                                    fsT.WriteText "\end_layout" & vbLf
                                    fsT.WriteText "" & vbLf
                                Else
                                    fsT.WriteText "\begin_layout Itemize" & vbLf
                                    fsT.WriteText ToItemize(txt) & vbLf
                                    fsT.WriteText "\end_layout" & vbLf
                                    fsT.WriteText "" & vbLf
                                End If
                            End If
                        End If
                    Next Pgh
                    If il > RealIndent(objshape.TextFrame.TextRange.Paragraphs.IndentLevel) Then
                      For i = 1 To il
                        fsT.WriteText "\end_deeper" & vbLf
                      Next i
                    End If
                End If
            ElseIf objshape.HasTable Then
              ln = "\begin{tabular}{|"
              For j = 1 To objshape.Table.Columns.Count
              ln = ln & "l|"
              Next j
              ln = ln & "} \hline"
              fsT.WriteText ln & vbLf
              With objshape.Table
                For i = 1 To .Rows.Count
                    If .Cell(i, 1).Shape.HasTextFrame Then
                        ln = .Cell(i, 1).Shape.TextFrame.TextRange.Text
                    End If

                    For j = 2 To .Columns.Count
                        If .Cell(i, j).Shape.HasTextFrame Then
                            ln = ln & " & " & .Cell(i, j).Shape.TextFrame.TextRange.Text
                        End If
                    Next j
                    ln = ln & "  \\ \hline"
                    fsT.WriteText ln & vbLf
                Next i
                fsT.WriteText "\end{tabular}" & vbCrLf & vbLf
              End With
            ElseIf (objshape.Type = msoGroup) Then
                For Each objGrpItem In objshape.GroupItems
                    If objGrpItem.HasTextFrame = True Then
                        If Not objGrpItem.TextFrame.TextRange Is Nothing Then
                           shpx = objGrpItem.Top / hght
                           shpy = objGrpItem.Left / wdth
                           If shpx < 0.1 And shpy > 0.5 Then
                            fsT.WriteText ("%BookTitle: " & objGrpItem.TextFrame.TextRange.Text) & vbLf
                            ElseIf shpx < 0.1 And shpy < 0.5 Then
                            fsT.WriteText ("%FrameTitle: " & objGrpItem.TextFrame.TextRange.Text) & vbLf
                            Else
                            fsT.WriteText ("%PartTitle: " & objGrpItem.TextFrame.TextRange.Text) & vbLf
                           End If
                        End If
                    End If
                 Next objGrpItem
            ElseIf (objshape.Type = msoPicture) Then
                IName = "img" & Format(ctr, "0000") & ".png"
                
                ' Are you kidding, VBA?
                currwidth = objshape.Width
                objshape.ScaleWidth 1, True
                orgwidth = objshape.Width
                scaleFactor = currwidth / orgwidth
                objshape.ScaleWidth scaleFactor, True
                
                fsT.WriteText "\begin_layout Standard" & vbLf
                fsT.WriteText "\begin_inset Graphics" & vbLf
                fsT.WriteText " filename " & IName & vbLf
                fsT.WriteText " lyxscale 20" & vbLf
                fsT.WriteText " scale " & PrintFloat(scaleFactor * 100) & "" & vbLf
                fsT.WriteText "" & vbLf
                fsT.WriteText "\end_inset" & vbLf
                fsT.WriteText "\end_layout" & vbLf
                fsT.WriteText "" & vbLf
                Call objshape.Export(Pth & "\" & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                ctr = ctr + 1
            ElseIf objshape.Type = msoEmbeddedOLEObject Then
                If objshape.OLEFormat.ProgID = "Equation.3" Then
                    IName = "img" & Format(ctr, "0000") & ".png"
                    fsT.WriteText "\begin_layout Standard" & vbLf
                    fsT.WriteText "\begin_inset Graphics" & vbLf
                    fsT.WriteText " filename " & IName & vbLf
                    fsT.WriteText " lyxscale 20" & vbLf
                    fsT.WriteText " scale " & PrintFloat(scaleFactor * 100) & "" & vbLf
                    fsT.WriteText "" & vbLf
                    fsT.WriteText "\end_inset" & vbLf
                    fsT.WriteText "\end_layout" & vbLf
                    fsT.WriteText "" & vbLf
                    Call objshape.Export(Pth & "\" & IName, ppShapeFormatPNG, , , ppRelativeToSlide)
                    ctr = ctr + 1
                End If
            Else
                fsT.WriteText objshape.Type & vbLf
            End If
        Next objshape


        fsT.WriteText "\end_deeper" & vbLf    ' End frame content
        fsT.WriteText "\begin_layout Standard" & vbLf
        fsT.WriteText "\begin_inset Separator plain" & vbLf  ' Frame separator
        fsT.WriteText "\end_inset" & vbLf
        fsT.WriteText "\end_layout" & vbLf
    Next objSlide
    
    fsT.WriteText "" & vbLf
    fsT.WriteText "\end_body" & vbLf
    fsT.WriteText "\end_document" & vbLf
    fsT.SaveToFile Dest, 2 'Save as binary
End Sub

Function ToItemize(Tex As String) As String
    Trim (Tex)
    sanitized = Replace(Tex, "<", "\begin_inset ERT " & vbLf & "status collapsed" & vbLf & vbLf & "\begin_layout Plain Layout" & vbLf & vbLf & "\backslash" & vbLf & "textless{}" & vbLf & "\end_layout" & vbLf & "\end_inset" & vbLf)
    ToItemize = Replace(sanitized, ">", "\begin_inset ERT " & vbLf & "status collapsed" & vbLf & vbLf & "\begin_layout Plain Layout" & vbLf & vbLf & "\backslash" & vbLf & "textgreater{}" & vbLf & "\end_layout" & vbLf & "\end_inset" & vbLf)
End Function

Function RealIndent(indent As Long) As Long
    If indent > 10 Or indent < 0 Then
        RealIndent = 1
    ElseIf indent > 2 Then
        RealIndent = 2   ' Latex itemize does not support more than three levels
    Else
        RealIndent = indent
    End If
End Function

Function PrintFloat(num As Variant) As String
    PrintFloat = Replace(Format(num, "##.00"), ",", ".")
End Function






