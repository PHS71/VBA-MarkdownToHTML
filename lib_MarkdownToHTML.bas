Attribute VB_Name = "lib_MarkdownToHTML"
Option Explicit

''
' @description
' Convertion d'une chaine de caractères en Markdown en son équivalent HTML
'
' @param {String} mdStr - Chaine en Markdown
' @returns {String} Chaine convertie en HTML
'
' @author PHS71
' @version 1 du 09/04/2017
'
Public Function MarkdownToHTML(ByVal mdStr As String) As String
On Error GoTo ErrorHandler

    'Titre
    If Left(mdStr, 2) = "# " Then
        MarkdownToHTML = "<h1>" & Mid(mdStr, 3) & "</h1>"
    ElseIf Left(mdStr, 3) = "## " Then
        MarkdownToHTML = "<h2>" & Mid(mdStr, 4) & "</h2>"
    ElseIf Left(mdStr, 4) = "### " Then
        MarkdownToHTML = "<h3>" & Mid(mdStr, 5) & "</h3>"
    ElseIf Left(mdStr, 5) = "#### " Then
        MarkdownToHTML = "<h4>" & Mid(mdStr, 6) & "</h4>"
    Else
        MarkdownToHTML = mdStr
    End If
    
Exit Function
ErrorHandler:
    VBA.err.Raise VBA.err.Number, "MarkdownToHTML", VBA.err.Description, VBA.err.HelpFile, VBA.err.HelpContext
End Function
