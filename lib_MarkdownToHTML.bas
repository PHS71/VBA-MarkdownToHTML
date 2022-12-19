Attribute VB_Name = "lib_MarkdownToHTML"
Option Explicit

''
' @description
' Converts a Markdown string to an HTML string
'
' @param {String} mdStr - Markdown string
' @returns {String} HTML string
'
' @author PHS71
' @version 2 - 2019-05-19
'
Public Function MarkdownToHTML(ByVal mdStr As String) As String
Const CodeName As String = "MarkdownToHTML"

On Error GoTo ErrorHandler

Const RE_HEADER_1 As String = "^<p># *(.+)<\/p>$"
Const RE_HEADER_2 As String = "^<p>## *(.+)<\/p>$"
Const RE_HEADER_3 As String = "^<p>### *(.+)<\/p>$"
Const RE_HEADER_4 As String = "^<p>#### *(.+)<\/p>$"
Const RE_HEADER_5 As String = "^<p>##### *(.+)<\/p>$"
Const RE_HEADER_6 As String = "^<p>###### *(.+)<\/p>$"

Const RE_LINEBREAK As String = " {2}<\/p>$"
Const RE_HORIZONTAL_RULE As String = "^<p>-{3,}<\/p>$"

Const RE_LIST_1 As String = "^<p>[\*|-] +(.+)<\/p>$"
Const RE_LIST_2 As String = "^<p>[0-9]+\. +(.+)<\/p>$"

Const RE_BLOCKQUOTE As String = "^<p>> *(.+)<\/p>$"

Const RE_CODE As String = "^<p> {4}(.*)<\/p>$"
Const RE_CODE_INLINE As String = "`([^`\n]+)`"

Const RE_STRONG_1 As String = "\*\*([^**\n]+)\*\*"
Const RE_STRONG_2 As String = "__([^__\n]+)__"

Const RE_EMPHASE_1 As String = "\*([^*\n]+)\*"
Const RE_EMPHASE_2 As String = "_([^_\n]+)_"

Const RE_IMAGE As String = "!\[([^]\n]+)\] *\(([^)\n""]+) ""([^""\n]+)""\)"
Const RE_LINK As String = "\[([^]\n]+)\] *\(([^)\n]+)\)"

Const ESC_UNDERSCORE As String = "\UNDERSCORE\CHAR" '_
Const ESC_STAR As String = "\STAR\CHAR" '*
Const ESC_HYPHEN As String = "\HYPHEN\CHAR" '-
Const ESC_BACKQUOTE As String = "\BACKQUOTE\CHAR" '`

Dim re As Object 'VBScript.RegExp

    Set re = VBA.CreateObject("VBScript.RegExp")
    
    re.Global = True
    re.MultiLine = True
    
    mdStr = "<p>" & mdStr & "</p>"
    mdStr = VBA.Replace(mdStr, VBA.vbCrLf, VBA.vbLf)
    mdStr = VBA.Replace(mdStr, VBA.vbLf, "</p>" & VBA.vbLf & "<p>")
    
    'Escape characters
    mdStr = VBA.Replace(mdStr, "\_", ESC_UNDERSCORE)
    mdStr = VBA.Replace(mdStr, "\*", ESC_STAR)
    mdStr = VBA.Replace(mdStr, "\-", ESC_HYPHEN)
    mdStr = VBA.Replace(mdStr, "\`", ESC_BACKQUOTE)
    
    'Headers
    re.Pattern = RE_HEADER_6
    mdStr = re.Replace(mdStr, "<h6>$1</h6>")
    re.Pattern = RE_HEADER_5
    mdStr = re.Replace(mdStr, "<h5>$1</h5>")
    re.Pattern = RE_HEADER_4
    mdStr = re.Replace(mdStr, "<h4>$1</h4>")
    re.Pattern = RE_HEADER_3
    mdStr = re.Replace(mdStr, "<h3>$1</h3>")
    re.Pattern = RE_HEADER_2
    mdStr = re.Replace(mdStr, "<h2>$1</h2>")
    re.Pattern = RE_HEADER_1
    mdStr = re.Replace(mdStr, "<h1>$1</h1>")
    
    'Code
    re.Pattern = RE_CODE
    mdStr = re.Replace(mdStr, "<pre>$1</pre>")
    mdStr = VBA.Replace(mdStr, "</pre>" & VBA.vbLf & "<pre>", VBA.vbLf)
    
    'BlockQuote
    re.Pattern = RE_BLOCKQUOTE
    mdStr = re.Replace(mdStr, "<blockquote>$1</blockquote>")
    mdStr = VBA.Replace(mdStr, "</blockquote>" & VBA.vbLf & "<blockquote>", VBA.vbLf)
    
    'Line break
    re.Pattern = RE_LINEBREAK
    mdStr = re.Replace(mdStr, "<br /></p>")
    
    'Horizontal Rule
    re.Pattern = RE_HORIZONTAL_RULE
    mdStr = re.Replace(mdStr, "<hr />")
    
    'Lists
    re.Pattern = RE_LIST_1
    mdStr = re.Replace(mdStr, "<ul><li>$1</li></ul>")
    mdStr = VBA.Replace(mdStr, "</ul>" & VBA.vbLf & "<ul>", VBA.vbLf)
    re.Pattern = RE_LIST_2
    mdStr = re.Replace(mdStr, "<ol><li>$1</li></ol>")
    mdStr = VBA.Replace(mdStr, "</ol>" & VBA.vbLf & "<ol>", VBA.vbLf)
    
    'Code in line
    re.Pattern = RE_CODE_INLINE
    mdStr = re.Replace(mdStr, "<code>$1</code>")
    
    'Strong
    re.Pattern = RE_STRONG_1
    mdStr = re.Replace(mdStr, "<strong>$1</strong>")
    re.Pattern = RE_STRONG_2
    mdStr = re.Replace(mdStr, "<strong>$1</strong>")
    
    'Emphase
    re.Pattern = RE_EMPHASE_1
    mdStr = re.Replace(mdStr, "<em>$1</em>")
    re.Pattern = RE_EMPHASE_2
    mdStr = re.Replace(mdStr, "<em>$1</em>")
    
    'Image
    re.Pattern = RE_IMAGE
    mdStr = re.Replace(mdStr, "<img alt=""$1"" title=""$3"" src=""$2"" />")
    
    'Link
    re.Pattern = RE_LINK
    mdStr = re.Replace(mdStr, "<a href=""$2"" target=""_blank"">$1</a>")
    
    'Escape characters
    mdStr = VBA.Replace(mdStr, ESC_UNDERSCORE, "_")
    mdStr = VBA.Replace(mdStr, ESC_STAR, "*")
    mdStr = VBA.Replace(mdStr, ESC_HYPHEN, "-")
    mdStr = VBA.Replace(mdStr, ESC_BACKQUOTE, "`")
    
    'Paragraphs
    mdStr = VBA.Replace(mdStr, VBA.vbLf & "<p></p>" & VBA.vbLf, VBA.vbLf & VBA.vbLf)
    mdStr = VBA.Replace(mdStr, "</p>" & VBA.vbLf & "<p>", VBA.vbLf)
    
    MarkdownToHTML = mdStr
    
    Set re = Nothing
    
Exit Function
ErrorHandler:
    VBA.Err.Raise VBA.Err.Number, CodeName, VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
End Function
