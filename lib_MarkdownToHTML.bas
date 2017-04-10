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
' @version 1.1 du 10/04/2017
'
Public Function MarkdownToHTML(ByVal mdStr As String) As String
On Error GoTo ErrorHandler

Dim reg As Object 'vbscript.regexp
Dim idx As Long

    Set reg = VBA.CreateObject("vbscript.regexp")
    
    'Titre : commence par #
    reg.Pattern = "^#* "
    If reg.test(mdStr) Then
        idx = InStr(1, mdStr, " ") - 1
        MarkdownToHTML = "<h" & CStr(idx) & ">" & Mid(mdStr, idx + 2) & "</h" & CStr(idx) & ">"
        Exit Function '-> on sort
    End If
    
    'Strong : **mon texte** ou __mon texte__
    reg.Pattern = "(.*)\*\*(.*)\*\*(.*)"
    mdStr = reg.Replace(mdStr, "$1<strong>$2</strong>$3")
    reg.Pattern = "(.*)__(.*)__(.*)"
    mdStr = reg.Replace(mdStr, "$1<strong>$2</strong>$3")
    
    'Emphase : *mon texte* ou _mon texte_
    reg.Pattern = "(.*)\*(.*)\*(.*)"
    mdStr = reg.Replace(mdStr, "$1<em>$2</em>$3")
    reg.Pattern = "(.*)_(.*)_(.*)"
    mdStr = reg.Replace(mdStr, "$1<em>$2</em>$3")
    
    MarkdownToHTML = mdStr
    
    Set reg = Nothing
    
Exit Function
ErrorHandler:
    VBA.Err.Raise VBA.Err.Number, "MarkdownToHTML", VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
End Function
