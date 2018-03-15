' refresh
' The macro is useful, if complex basic-files are used to build geometry.
'
' Different to standard structure macros (*.mcs, where only one subroutine "sub main" is allowed),
' those included bas-files MAY contain additional subroutines and functions.
'
' Existing parameters can be used in the bas-file without any previous definition,
' which makes this solution even accessable during parameter sweeping and optimization.
'
'------------------------------------------------------------------------------------------
' 23-Oct-2015 ube: cancel is now working properly
' 11-Mar-2008 ube: first version
'------------------------------------------------------------------------------------------

'Option Explicit

Sub Main ()
ResetAll
Dim sfilename As String
Dim A As Variant
Dim i As Integer

A = Array("CMC_dialogV3.vba","cmc_drawingV3.vba","freq_dom.vba")

' beginHide
For i = LBound(A)  To UBound(A)
If Dir("Z:\CST\Macros\Passive\CMC\"+(A(i))) <> "" Then
	Debug.Print(i)
	Debug.Print(A(i))

RunScript("Z:\CST\Macros\Passive\CMC\"+A(i))
Else
MsgBox "not Existing"
Exit All
End If
Next i
' endHide
'RunScript("Z:\CST\Macros\Passive\CMC\CMC_DialogV3.vba")
'RunScript("Z:\CST\Macros\Passive\CMC\cmc_drawingV3.vba")
'RunScript("C:\Users\MoonenDJG1\AppData\Roaming\CST AG\CST STUDIO SUITE\Library\Macros\freq_dom.mcs")

End Sub

