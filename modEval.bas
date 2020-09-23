Attribute VB_Name = "modEval"
''''''''''''''''''''''''''''
' semi's eval              '
' 3rd try                  '
' stefan@seemayer.de       '
' www.semicolonsoftware.de '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' calculates the result of an expression, e.g. 2*(3+4+5)^2
' correctly (i hope *g*).
'
' operators allowed:
' + add
' - substract
' * multiply
' / divide
' ^ potentiate
'
' requirements
' vb
'
' this was coded after my pc fucked up and i had nothing to do.
' i installed vb6 on my 133mhz crap laptop and wanted to code
' something useful, so i remembered that i had tried coding this
' a thousand times already.n
' since i had nothing else to do, i tried one more time, with my
' fingers freezing (cold german winters *g*, living in an old house)
' and succeeded.
'
' this was coded 100% by myself, using pure vb, while i was freezing,
' using no additional libraries, and i'm quite proud now :)
'
' post questions, bug reports and comments at the entry at pscode.com
'
' sorry for my bad english, i used some german variable names in some
' routines, but that shouldn't be too bad ;)
'
' License
' -------
' You may use this code in any free application
' as long as you give credit and link to http://www.semicolonsoftware.de
'
' If you would like to use it in a commercial application,
' please contact me first at stefan@seemayer.de




''''''''''''''''''
' and pls vote!!!'
''''''''''''''''''

Public Function EvalSimpleExpression(e As String) As Double
'Evaluates a simple expression,
'e.g. 2+3




Dim rz As String
Dim links As Double
Dim rechts As Double
Dim ergebnis As Double

If InStr(1, e, "+") > 0 Then rz = "+"
If InStr(1, e, "-") > 0 Then rz = "-"
If InStr(1, e, "*") > 0 Then rz = "*"
If InStr(1, e, "/") > 0 Then rz = "/"
If InStr(1, e, "^") > 0 Then rz = "^"


Dim rzpos As Integer
rzpos = InStr(1, e, rz)
links = Left(e, rzpos - 1)
rechts = Mid(e, rzpos + 1)

Select Case rz
Case "+"
ergebnis = links + rechts
Case "-"
ergebnis = links - rechts
Case "*"
ergebnis = links * rechts
Case "/"
ergebnis = links / rechts
Case "^"
ergebnis = links ^ rechts
End Select

EvalSimpleExpression = ergebnis

End Function

Public Function BracketsForExpression(e As String) As String
Dim areas As New Collection
Dim operators As New Collection


Dim exp As String

'remove spaces
exp = Replace(e, " ", "")

Dim j As Long

Dim c As String
Dim brk As String

Dim lastarea As String

'this bit will go through the expression, splitting it
'into different parts.
'
'all numbers and expressions will be saved into areas,
'the operators will be saved into operators

Do While i <= Len(exp)
i = i + 1
c = Mid(exp & " ", i, 1)

'Is digit?
If IsNumeric(c) Then
'go on searching
lastarea = lastarea & c
ElseIf c = "(" Then ' bracket handling
j = InStrRev(exp, ")")
brk = Mid(exp, i + 1, j - i - 1) 'cut out bracket
brk = BracketsForExpression(brk) 'add extra brackets to expression
lastarea = lastarea & brk 'add to lastarea
i = j 'continue at end of bracket
Else
areas.Add lastarea 'one block has ended, save number
operators.Add c    'save operator
lastarea = ""      'start searching again

End If

Loop

Dim op As Integer
Dim ergb
Do


If areas.Count = 1 Then Exit Do ' loop until all areas are transformed into one expression
op = GetNextOperator(operators) ' get next operator to work with (* before +, ^ before *)

ergb = "(" & areas(op) & operators(op) & areas(op + 1) & ")" ' build expression from collection
areas.Remove op + 1 'remove old values
areas.Remove op
operators.Remove op 'remove old operator

'add new expression
If areas.Count = 0 Or op > areas.Count Then
areas.Add ergb
Else
areas.Add ergb, , op
End If

'log created expression
l CStr(ergb)

Loop

' return value: all array entries transformed into one expression
BracketsForExpression = areas(1)

End Function

Public Function GetNextOperator(ops As Collection) As Integer
' Finds out which operator will be next

Dim highestoplevel As Long
Dim highestop As Long
For i = 1 To ops.Count
    If GetOpLevel(ops(i)) > highestoplevel Then
        highestop = i
        ho = ops(highestop)
        highestoplevel = GetOpLevel(ops(i))
    End If
Next i

GetNextOperator = highestop
End Function

Public Function IsNumeric(c As String) As Boolean
'is it a digit?
IsNumeric = (InStr(1, "0123456789,.", c) > 0)
End Function

Public Function GetOpLevel(o As String) As Integer
' get operator importance
Select Case o
Case "+"
GetOpLevel = 1
Case "-"
GetOpLevel = 1
Case "*"
GetOpLevel = 2
Case "/"
GetOpLevel = 2
Case "^"
GetOpLevel = 3
End Select
End Function


Public Function EvalBrackets(e As String) As Double
'evaluates bracket expression

Dim exp As String
'remove spaces
exp = Replace(e, " ", "")

Dim hblpos As Long
Dim hblposend As Long

Dim hb As String
Dim hbe As String

'loop until there are no more brackets
Do Until InStr(1, exp, "(") = 0

'first calculate the bracket with highest level
hblpos = GetHighestBracketLevelPos(exp)
hblposend = InStr(hblpos, exp, ")")

'cut out bracket
hb = Mid(exp, hblpos + 1, (hblposend - hblpos) - 1)

'eval it's expression
hbe = EvalSimpleExpression(hb)

'replace with results
exp = Replace(exp, "(" & hb & ")", hbe)

Loop

'return value = our result
EvalBrackets = exp

End Function

Public Function GetHighestBracketLevelPos(e As String) As Long
'finds out which bracket has the highest level
'and returns position in string

Dim hbl As Long
Dim hblpos As Long
Dim i As Long

For i = 1 To Len(e)
If GetBracketLevelAt(e, i) > hbl Then
hbl = GetBracketLevelAt(e, i)
hblpos = i
End If
Next i

GetHighestBracketLevelPos = hblpos

End Function

Public Function GetBracketLevelAt(e As String, pos As Long) As Integer
'gets the level of the bracket at position pos
Dim bl As Integer
For i = 1 To pos
    If Mid(e, i, 1) = "(" Then bl = bl + 1
    If Mid(e, i, 1) = ")" Then bl = bl - 1
Next i
GetBracketLevelAt = bl
End Function

Public Function Eval(Expression As String) As Double
'eval function,
'this is what you will call from your app!

Dim exp As String
'remove spaces
exp = Replace(Expression, " ", "")

'puts brackets so * is calculated before + etc.
exp = BracketsForExpression(exp)

'evaluates the expression
Eval = EvalBrackets(exp)
End Function

Public Sub l(t As String)
' add logging code here!

'debug.Print t          'log to debug window
'Form1.log.AddItem t    'log to listbox
End Sub
