Option Explicit On 
Option Strict On

Public Interface ICutCopyPaste
    Sub Cut()
    Sub Copy()
    Sub Paste()
    Sub SelectAll()
End Interface

Public Interface IFind
    Sub Find()
    Sub FindNext()
End Interface
