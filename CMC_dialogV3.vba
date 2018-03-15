' CMC_dialog

Sub Main ()
' msgBox "start subroutine"
if GetNumberOfParameters <> 0 Then
msgBox "No parameters, exiting subroutine"
exit sub	
end If 	

'load all parameters into script, they originate from CMC_dialogV3'
' For kk = 0 To GetNumberOfParameters-1
' 	mystr = GetParameterName(kk)
' 	num = GetParameterNValue(kk)
' 	Debug.Print((mystr + " = " + CStr(num)))
' 	Debug.Print(Evaluate(mystr))
'   MakeSureParameterExists(mystr,RestoreParameter(mystr))
' Next kk

Dim cst_core_r As Double, cst_core_w As Double, cst_core_h As Double
Dim cst_core_x As Double, cst_core_y As Double, cst_core_z As Double, cst_wire_r As Double
Dim cst_core_ang As Double, cst_core_off As Double
Dim cst_lead As Double, cst_phases_N As Double
Dim cst_torrus_x As Double, cst_torrus_y As Double, cst_torrus_z As Double
Dim cst_core_ra As Double, cst_core_ri As Double
Dim cst_wire_N As Integer, cst_result As Integer, cst_symm_term As Integer, cst_torrus_i As Integer
Dim phases_i As Integer, cst_kern As Integer,  cst_simp As Integer, cst_core_i As Integer
Dim scst_core_ri As String, scst_core_ra As String, scst_core_h As String, scst_phases As String, scst_lead As String
Dim scst_core_x As String, scst_core_y As String, scst_core_z As String, scst_wire_r As String, scst_kern As String
Dim scst_wire_N As String, scst_symm_term As Integer, scst_core_ang As String, scst_core_off As String
Dim scst_h_gnd As String, cst_h_gnd As Double
Dim AA As Variant

AA = Array("textVals","WE-S-744822-301","WE-S-744822-222","WE-S-744822-233","WE-S-744822-110","WE-S-744822-120","WE-M-744823-601","WE-M-744823-422","WE-M-744823-305","WE-M-744823-210","WE-M-744823-220","Vitroperm - X363")

	Begin Dialog preSets 400,203 ' %GRID:10,7,1,1
		DropListBox 20,14,300,49,AA,.DropListBox1
		OKButton 30,84,90,21
		CancelButton 170,84,90,21
	End Dialog
	Dim pre As preSets

	' cst_result2 As Integer
	cst_result2=Dialog(pre)
	msgBox CStr(cst_result2)
	' cst_result2 = Dialog(pre)
	BeginHide
    assign "cst_result2"
	EndHide
  	If (cst_result2 = 0) Then Exit sub



'HAND1:
	Begin Dialog UserDialog 640,343,"Create 3D Toroidal Coil, rectangular core." ' %GRID:10,7,1,1
		Text 30,25,120,14,"Core inner radius",.Text1
		Text 30,55,130,14,"Core outer radius",.Text2
		Text 30,85,90,14,"Core height",.Text3
		Text 30,115,90,14,"Wire radius",.Text4
		Text 30,145,90,14,"Turns",.Text5
		Text 30,175,90,14,"Angle in rad",.Text6
		Text 30,205,90,14,"Angle Offset",.Text7
		Text 30,235,90,14,"Lead",.Text8
		Text 30,265,90,14,"N-Phases",.Text9
		Text 30,295,90,14,"height_gnd",.Text10
		TextBox 170,20,90,21,.ri
		TextBox 170,50,90,21,.ra
		TextBox 170,80,90,21,.h
		TextBox 170,110,90,21,.wr
		TextBox 170,140,90,21,.n
		TextBox 170,170,90,21,.ang
		TextBox 170,200,90,21,.off
		TextBox 170,230,90,21,.ld
		TextBox 170,260,90,21,.ph
		TextBox 170,290,90,21,.h_gnd
		OptionGroup .option_kern
			OptionButton 310,210,160,21,"Core Off",.option_kern_off
			OptionButton 310,231,190,14,"Core On",.option_kern_on
		OptionGroup .option_simp
			OptionButton 310,251,160,21,"Simplify Off",.option_simp_off
			OptionButton 310,271,190,14,"Simplify On",.option_simp_on
		OptionGroup .option_ports
			OptionButton 310,290,160,21,"2-Port",.option_ports_2
			OptionButton 310,311,190,14,"4-Port",.option_ports_4
		Picture 290,7,340,168,GetInstallPath + "\Library\Macros\Construct\Coils\3D Toroidal Coil - rectangular core.bmp",0,.Picture1
		OKButton 30,312,90,21
		CancelButton 130,312,90,21
		End Dialog
	Dim dlg As UserDialog
	' If (cst_result = 0) Then Exit All
	'Default values
	
	If pre.DropListBox1=0 Then
	dlg.ri = "0"
	dlg.ra = "0"
	dlg.h = "0"
	dlg.wr = "0"
	dlg.n = "0"
	dlg.ang = "0"
	dlg.off = "0"
	dlg.ld = "0"
	dlg.ph = "0"
	dlg.h_gnd = "0"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=1 Then
	dlg.ri = "3.6"
	dlg.ra = "7.6"
	dlg.h = "9.2"
	dlg.wr = "0.2"
	dlg.n = "12"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=2 Then
	dlg.ri = "3.6"
	dlg.ra = "7.6"
	dlg.h = "9.2"
	dlg.wr = "0.2"
	dlg.n = "17"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=3 Then
	dlg.ri = "3.6"
	dlg.ra = "7.6"
	dlg.h = "9.2"
	dlg.wr = "0.2"
	dlg.n = "21"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=4 Then
	dlg.ri = "3.6"
	dlg.ra = "7.6"
	dlg.h = "9.2"
	dlg.wr = "0.2"
	dlg.n = "35"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=5 Then
	dlg.ri = "3.6"
	dlg.ra = "7.6"
	dlg.h = "9.2"
	dlg.wr = "0.2"
	dlg.n = "43"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=6 Then
	dlg.ri = "4.55"
	dlg.ra = "8.2"
	dlg.h = "10.4"
	dlg.wr = "0.2"
	dlg.n = "10"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=7 Then
	dlg.ri = "4.55"
	dlg.ra = "8.2"
	dlg.h = "10.4"
	dlg.wr = "0.2"
	dlg.n = "16"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=8 Then
	dlg.ri = "4.55"
	dlg.ra = "8.2"
	dlg.h = "10.4"
	dlg.wr = "0.2"
	dlg.n = "26"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=9 Then
	dlg.ri = "4.55"
	dlg.ra = "8.2"
	dlg.h = "10.4"
	dlg.wr = "0.2"
	dlg.n = "34"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=10 Then
	dlg.ri = "4.55"
	dlg.ra = "8.2"
	dlg.h = "10.4"
	dlg.wr = "0.2"
	dlg.n = "45"
	dlg.ang = "2.2"
	dlg.off = "2"
	dlg.ld = "2"
	dlg.ph = "2"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	ElseIf pre.DropListBox1=11 Then
	dlg.ri = "17.725"
	dlg.ra = "26.75"
	dlg.h = "23.3"
	dlg.wr = "0.5"
	dlg.n = "14"
	dlg.ang = "1.38"
	dlg.off = "1.57"
	dlg.ld = "2"
	dlg.ph = "3"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	Else
	dlg.ri = "17.725"
	dlg.ra = "26.75"
	dlg.h = "23.3"
	dlg.wr = "0.5"
	dlg.n = "14"
	dlg.ang = "1.38"
	dlg.off = "1.57"
	dlg.ld = "2"
	dlg.ph = "3"
	dlg.h_gnd = "0.5"
	dlg.option_kern = 1
	dlg.option_simp = 0
	End If
	dlg.option_ports = 0

	cst_result = Dialog(dlg)
BeginHide
    assign "cst_result"
EndHide
  If (cst_result = 0) Then Exit All
  scst_core_ri = dlg.ri  ' Core radius
  scst_core_ra = dlg.ra	' core width
  scst_core_h = dlg.h	' core height
  scst_wire_r = dlg.wr	' wire radius
  scst_wire_N = dlg.n	' number of turns
  scst_core_ang = dlg.ang ' angle of windings
  scst_core_off = dlg.off ' offset of angle
  scst_lead = dlg.ld
  scst_phases = dlg.ph
  scst_h_gnd = dlg.h_gnd
  cst_kern = Cint(dlg.option_kern)
  cst_simp = Cint(dlg.option_simp)


BeginHide
  assign "scst_core_ri"       ' writes e.g. "cst_core_ri = 0.1"     into history list
  assign "scst_core_ra"
  assign "scst_core_h"
  assign "scst_wire_r"
  assign "scst_wire_N"
  assign "scst_core_ang"
  assign "scst_core_off"
  assign "scst_symm_term"
  assign "scst_lead"
  assign "scst_phases"
  assign "cst_kern"
  assign "cst_simp"
  assign "scst_h_gnd"
  assign "cst_wire_r"
  ' assign "cst_phases_N"
EndHide

  

Debug.Print(pre.DropListBox1)

        'cst_result = Evaluate(cst_result)
        If (cst_result =0) Then Exit All   ' if cancel/help is clicked, exit all
        If (cst_result =1) Then Exit All
        cst_core_r       = 0.5 * (Evaluate(scst_core_ri) + Evaluate(scst_core_ra))
		cst_core_w       = Evaluate(scst_core_ra) - Evaluate(scst_core_ri)
		cst_core_h       = Evaluate(scst_core_h)
		cst_wire_r       = Evaluate(scst_wire_r)
		cst_wire_N       = Evaluate(scst_wire_N)
		cst_core_ang	 = Evaluate(scst_core_ang)
		cst_core_off	 = Evaluate(scst_core_off)
		cst_lead	 = Evaluate(scst_lead)
		cst_symm_term = cint(Evaluate(scst_symm_term))
		cst_core_ra = Evaluate(scst_core_ra)
		cst_core_ri = Evaluate(scst_core_ri)
		cst_phases_N = Evaluate(scst_phases)
		cst_h_gnd = Evaluate(scst_h_gnd)

REM BeginHide
	StoreDoubleParameter "cst_core_r", (Evaluate(scst_core_ri) + Evaluate(scst_core_ra))
	SetParameterDescription  ( "cst_core_r",  "radius of the core middle"  )

	StoreDoubleParameter "cst_wire_r", Evaluate(scst_wire_r)
	SetParameterDescription  ( "cst_wire_r",  "radius of the wire"  )

	StoreDoubleParameter "cst_core_w", Evaluate(scst_core_ra) - Evaluate(scst_core_ri)
	SetParameterDescription  ( "cst_core_w",  "Core Width"  )

	StoreDoubleParameter "cst_core_h", Evaluate(scst_core_h)
	SetParameterDescription  ( "cst_core_h",  "Core Height"  )

	StoreDoubleParameter "cst_wire_N", Evaluate(scst_wire_N)
	SetParameterDescription  ( "cst_wire_N",  "Amount of Windings"  )

	StoreDoubleParameter "cst_core_ang", Evaluate(scst_core_ang)
	SetParameterDescription  ( "cst_core_ang",  "Angle of winding"  )

	StoreDoubleParameter "cst_core_off", Evaluate(scst_core_off)
	SetParameterDescription  ( "cst_core_off",  "Angle offset of windings"  )

	StoreDoubleParameter "cst_lead",  Evaluate(scst_lead)
	SetParameterDescription  ( "cst_lead",  "Lead length"  )

	StoreDoubleParameter "cst_symm_term", cint(Evaluate(scst_symm_term))
	SetParameterDescription  ( "cst_symm_term",  "Symmetry"  )

	StoreDoubleParameter "cst_core_ra",  Evaluate(scst_core_ra)
	SetParameterDescription  ( "cst_core_ra",  "Core outer radius"  )

	StoreDoubleParameter "cst_core_ri", Evaluate(scst_core_ri)
	SetParameterDescription  ( "cst_core_ri",  "Core inner radius"  )


	StoreDoubleParameter "cst_phases_N", Evaluate(scst_phases)
	SetParameterDescription  ( "cst_phases_N",  "Amount of Phases"  )


	StoreDoubleParameter "cst_h_gnd", Evaluate(scst_h_gnd)
	SetParameterDescription  ( "cst_h_gnd",  "Height towards GND-plane"  )


	StoreDoubleParameter "cst_kern", Cint(dlg.option_kern)
	SetParameterDescription  ( "cst_kern",  "Option to not draw core"  )


	StoreDoubleParameter "cst_simp", Cint(dlg.option_simp)
	SetParameterDescription  ( "cst_simp",  "Option to simplify"  )

	StoreDoubleParameter "remove", Cint(0)
	SetParameterDescription  ( "remove",  "Option to Remove all parameters"  )



REM EndHide

End Sub
