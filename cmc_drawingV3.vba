'-----------------------------------------------------------------------------------------------------------------------------
' History Of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 15-Mar-2018 Niek Moonen: Created full vba script, drawing CMC above a ground-plane with variable height,turns,size etc.
' 07-Mar-2017 Niek Moonen: Add Simplified option
' 11-Jan-2016 Niek Moonen: Adding beginning and end angle of the windings.
' 05-Feb-2014 ube/fwe: dialog changes: inner and outer radius more intuitive, repaired picture, symmetric terminals always on
' 02-Jul-2012 fhi: symmetrical terminals at outer side of the windings
' 27-Apr-2012 fhi: option for wire-radius= 0 (creates curves), correcting initial point for curve
' 29-Sep-2011 ube: include picture, allow multiple execution for multiple coils
' 28-Sep-2011 jwa: hide dialog
' 18-Jan-2011 fwe: Initial version
'-----------------------------------------------------------------------------------------------------------------------------
'#Language "WWB-COM"

Option Explicit
Sub Main
'ResetAll
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
Dim mystr As String
Dim num As Double
Dim kk As Integer
' msgBox "start"
'load all parameters into script, they originate from CMC_dialogV3'
For kk = 0 To GetNumberOfParameters-1
	mystr = GetParameterName(kk)
	num = GetParameterNValue(kk)
	Debug.Print((mystr + " = " + CStr(num)))
	Debug.Print(Evaluate(mystr))
  MakeSureParameterExists(mystr,RestoreParameter(mystr))
Next kk

    cst_core_ra    = Evaluate("cst_core_ra")
    cst_core_ri    = Evaluate("cst_core_ri")
  	cst_core_r       = 0.5 * (Evaluate("cst_core_ra")+Evaluate("cst_core_ri"))
		cst_core_w       = Evaluate("cst_core_w")
		cst_core_h       = Evaluate("cst_core_h")
		cst_wire_r       = Evaluate("cst_wire_r")
		cst_wire_N       = Evaluate("cst_wire_N")
		cst_core_ang	 = Evaluate("cst_core_ang")
		cst_core_off	 = Evaluate("cst_core_off")
		cst_lead	 	 = Evaluate("cst_lead")
		'cst_symm_term 	 = (Evaluate("cst_symm_term")))

		cst_phases_N 	 = Evaluate("cst_phases_N")
		cst_h_gnd 		 = Evaluate("cst_h_gnd")


On Error Resume Next
 Curve.DeleteCurve "core_curve"
 Curve.DeleteCurve "wire_crosssection"
 Curve.DeleteCurve "torrus_curve"
On Error GoTo 0


'@ define boundaries
With Boundary
     .Xmin "electric"
     .Xmax "electric"
     .Ymin "electric"
     .Ymax "electric"
     .Zmin "electric"
     .Zmax "electric"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
End With


'-----------------------------------------------------------------------------------------------------------------------------
'Core drawing

WCS.ActivateWCS "global"
Debug.Print("start drawing core")
Debug.Print(Evaluate("cst_kern"))
Debug.Print("start drawing core")
If Evaluate("cst_kern") = 1 Then 'check if core needs to be created
	 'start with core creation
	 Curve.NewCurve "torrus_curve"
	 Debug.Print("start drawing core")

With Layer
     .Reset
     .Name "Ferrite"
     .FrqType "hf"
     .Type "Normal"
     .Epsilon "1.0"
     .Mue "1.0"
     .Kappa "0.0"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstTanD"
     .KappaM "0.0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .DispModelEps "None"
     .DispModelMue "None"
     .Rho "0.0"
     .Colour "0.501961", "0.501961", "0.501961"
     .Wireframe "False"
     .Transparency "0"
     .Create
 End With

  With Polygon3D
     .Reset
     .Name "torrus_3dpolygon"
     .Curve "torrus_curve"
     For cst_torrus_i = 0 To 360 'full circle!
     	cst_torrus_x = Evaluate("cst_core_r") *Cos(cst_torrus_i*pi/180)
     	cst_torrus_y = Evaluate("cst_core_r")*Sin(cst_torrus_i*pi/180)
     	cst_torrus_z = 0
      .Point cst_torrus_x, cst_torrus_y, cst_torrus_z
     Next cst_torrus_i
     .Create
 End With

Curve.NewCurve "torrus_crosssection"




With Rectangle
     .Reset
     .Name "rect1"
     .Curve "torrus_crosssection"
     .Xrange Evaluate("cst_core_ri") +Evaluate("cst_wire_r"), Evaluate("cst_core_ra")-Evaluate("cst_wire_r")
     .Yrange -Evaluate("cst_core_h")/2+Evaluate("cst_wire_r"), Evaluate("cst_core_h")/2-Evaluate("cst_wire_r")
     .Create
End With

With Transform
     .Reset
     .Name "torrus_crosssection:rect1"
     .Origin "Free"
     .Center "0", "0", "0"
     .Angle "90", "0", "0"
     .MultipleObjects "False"
     .GroupObjects "False"
     .Repetitions "1"
     .MultipleSelection "False"
     .Transform "Curve", "Rotate"
End With


With SweepCurve
     .Reset
     .Name Solid.GetNextFreeName
     .Component "Ferrite"
     .Material "Ferrite"
     .Twistangle "0.0"
     .Taperangle "0.0"
     .ProjectProfileToPathAdvanced "True"
     .Path "torrus_curve:torrus_3dpolygon"
     .Curve "torrus_crosssection:rect1"
     .Create
End With


End If
'-----------------------------------------------------------------------------------------------------------------------------



'-----------------------------------------------------------------------------------------------------------------------------
If cst_simp= 1 Then

For phases_i = 1 To Evaluate("cst_phases_N")

	If phases_i = 2 Then
		cst_core_off = Evaluate("cst_core_off") + 0.66*pi
	ElseIf phases_i = 3 Then
		cst_core_off = Evaluate("cst_core_off") + 0.66*pi

	End If

'only draw first winding
With Polygon3D
     .Reset
     .Name "core_3dpolygon"
     .Curve "core_curve"

     cst_core_i = 0

     cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
      .Point cst_core_x, cst_core_y, cst_core_z



      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=0.5*cst_core_h+cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z


	'add leads
     cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
     .Point cst_core_x, cst_core_y, cst_core_z

     .Create
     End With

'end draw first winding




'@ new curve: wire_crosssection
If cst_wire_r > 0 Then
	Curve.NewCurve "wire_crosssection"
End If

'@ store picked point: 1
Pick.NextPickToDatabase "1"
Pick.PickCurveEndpointFromId "core_curve:core_3dpolygon", "1"

If cst_wire_r > 0 Then

'@ define curve circle: wire_crosssection:circle1
With Circle
     .Reset
     .Name "circle1"
     .Curve "wire_crosssection"
     .Radius cst_wire_r
     .Xcenter cst_core_x 'cst_core_r+0.5*cst_core_w+cst_wire_r
     .Ycenter cst_core_y '"0"
     .Segments "0"
     .Create
End With

 With Transform
     .Reset
     .Name "wire_crosssection:circle1"
     .Origin "ShapeCenter"
     .Center "0", "0", "0"
     .Angle "0", "90", "0"
     .MultipleObjects "False"
     .GroupObjects "False"
     .Repetitions "1"
     .MultipleSelection "False"
     .Transform "Curve", "Rotate"
 End With

 With Transform
     .Reset
     .Name "wire_crosssection:circle1"
     .Vector "0", "0", cst_core_h/2+cst_wire_r
     .UsePickedPoints "False"
     .InvertPickedPoints "False"
     .MultipleObjects "False"
     .GroupObjects "False"
     .Repetitions "1"
     .MultipleSelection "False"
     .Transform "Curve", "Translate"
 End With



With Material
	If Not .Exists("Wire_material") Then
		.Reset
		.Name "Wire_material"
		.FrqType "hf"
		.Type "Pec"
		.Rho "0.0"
		.Colour "1", "0.501961", "0"
		.Wireframe "False"
		.Transparency "0"
		.Reflection "True"
		.Create
	End If
End With


'@ define sweepprofile: core:wire
With SweepCurve
     .Reset
     .Name Solid.GetNextFreeName
     .Component "Wire_material"
     .Material "Wire_material"
     .Twistangle "0.0"
     .Taperangle "0.0"
     .ProjectProfileToPathAdvanced "True"
     .Path "core_curve:core_3dpolygon"
     .Curve "wire_crosssection:circle1"
     .Create
End With

End If
Next phases_i
End If
' msgBox "geen idee"
If cst_simp = 0 Then 'non-simplification
For phases_i = 1 To cst_phases_N

	If cst_phases_N = 2 Then
		'cst_core_off = cst_core_off + 0.66*pi
		if phases_i = 2 then
				cst_core_off = cst_core_off - 1*0.33*pi
				end if
	ElseIf cst_phases_N = 38 Then
		cst_core_off = cst_core_off + 0.66*pi
	End If

With Polygon3D
     .Reset
     .Name "core_3dpolygon"
     .Curve "core_curve"

     cst_core_i = 0

'- opposite way of winding, often in CMC
If (cst_phases_N = 2 And phases_i = 2) Then
	'---------------------------------------- 1st winding
	
	
     cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
      .Point cst_core_x, cst_core_y, cst_core_z



      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=0.5*cst_core_h+cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

'------------------------------------------



For cst_core_i = 1 To cst_wire_N-1

      'cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang) 'original
      'cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang) 'original
      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_z=0.5*cst_core_h+cst_wire_r
      .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=0.5*cst_core_h+cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

     If cst_core_i = cst_wire_N-1 Then
     cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin(-(cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
     .Point cst_core_x, cst_core_y, cst_core_z
     End If

   Next cst_core_i

   cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
   cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin(-(cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
   cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
   Pick.NextPickToDatabase(1)
    .Create

    '- normal way of winding, all same direction
Else
	'---------------------------------------- 1st winding

     cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
      .Point cst_core_x, cst_core_y, cst_core_z



      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=0.5*cst_core_h+cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

'------------------------------------------
     For cst_core_i = 1 To cst_wire_N-1
      'cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang) 'original
      'cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang) 'original
      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_z=0.5*cst_core_h+cst_wire_r
      .Point cst_core_x, cst_core_y, cst_core_z



      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=0.5*cst_core_h+cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

     If cst_core_i = cst_wire_N-1 Then
     cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
     .Point cst_core_x, cst_core_y, cst_core_z
     End If

   Next cst_core_i

   cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
   cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
   cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
   Pick.NextPickToDatabase(1)
    .Create
End If



End With


'@ new curve: wire_crosssection
If cst_wire_r > 0 Then
	Curve.NewCurve "wire_crosssection"
End If


'@ store picked point: 1
Pick.PickCurveEndpointFromId("core_curve:core_3dpolygon",1)
WCS.SetOrigin(xp(1),yp(1),zp(1))
WCS.Store("End_of_wires"+CStr(phases_i))
'Debug.Print("kak")

If cst_wire_r > 0 Then

'@ define curve circle: wire_crosssection:circle1

With Circle
     .Reset
     .Name "circle1"
     .Curve "wire_crosssection"
     .Radius cst_wire_r
     .Xcenter cst_core_x 'cst_core_r+0.5*cst_core_w+cst_wire_r
     .Ycenter cst_core_y '"0"
     .Segments "0"
     .Create
End With

 With Transform
     .Reset
     .Name "wire_crosssection:circle1"
     .Origin "ShapeCenter"
     .Center "0", "0", "0"
     .Angle "0", "90", "0"
     .MultipleObjects "False"
     .GroupObjects "False"
     .Repetitions "1"
     .MultipleSelection "False"
     .Transform "Curve", "Rotate"
 End With

 With Transform
     .Reset
     .Name "wire_crosssection:circle1"
     .Vector "0", "0", cst_core_h/2+cst_wire_r
     .UsePickedPoints "False"
     .InvertPickedPoints "False"
     .MultipleObjects "False"
     .GroupObjects "False"
     .Repetitions "1"
     .MultipleSelection "False"
     .Transform "Curve", "Translate"
 End With



With Material
	If Not .Exists("Wire_material") Then
		.Reset
		.Name "Wire_material"
		.FrqType "hf"
		.Type "Pec"
		.Rho "0.0"
		.Colour "1", "0.501961", "0"
		.Wireframe "False"
		.Transparency "0"
		.Reflection "True"
		.Create
	End If
End With


'@ define sweepprofile: core:wire

With SweepCurve
     .Reset
     .Name Solid.GetNextFreeName
     .Component "Wire_material"
     .Material "Wire_material"
     .Twistangle "0.0"
     .Taperangle "0.0"
     .ProjectProfileToPathAdvanced "True"
     .Path "core_curve:core_3dpolygon"
     .Curve "wire_crosssection:circle1"
     .Create
End With

End If
Next phases_i
End If





' Curve.DeleteCurve "core_curve"
' Curve.DeleteCurve "wire_crosssection"
 If cst_kern = 1 Then
 Curve.DeleteCurve "torrus_curve"
 Curve.DeleteCurve "torrus_crosssection"
 End If
'-----------------------------------------------------------------------------------------------------------------------------
WCS.ActivateWCS "local"
WCS.SetOrigin(0,0,0)
WCS.Store "centre"

WCS.SetOrigin(0,0,0.5*cst_core_h+cst_lead+1)
WCS.Store "high"

drawlead(0,cst_h_gnd,cst_wire_r)

Dim gnd_x As Double
Dim gnd_y As Double
Dim gnd_z As Double
gnd_x=150
gnd_y=150
gnd_z=1

drawGND(cst_h_gnd,gnd_x,gnd_y,gnd_z)

ports()
setup()

End Sub

Sub drawlead(ByVal cst_simp As Boolean,ByVal cst_h_gnd As Double,ByVal cst_wire_r As Double)
Dim wireRad As Double
Dim wireLength As Double
wireLength=0.1
If cst_simp = 0 Then
'adding rect leads for ports
'@ pick face
If cst_phases_N=3 Then
Pick.PickFaceFromId "Wire_material:solid4", "2"
WCS.AlignWCSWithSelected "Face"
drawBrick_lead(cst_wire_r,wireLength)
WCS.Restore("End_of_wires3")
'Pick.PickFaceFromId "Wire_material:solid4", "57" 'only works for non-simplification
drawBrick_lead(cst_wire_r,wireLength)
End If


'@ pick face
Pick.PickFaceFromId "Wire_material:solid3", "2"
WCS.AlignWCSWithSelected "Face"
drawBrick_lead(cst_wire_r,wireLength)
'@ pick face
'Pick.PickFaceFromId "Wire_material:solid3", "57"
WCS.Restore("End_of_wires1")
drawBrick_lead(cst_wire_r,wireLength)
'@ pick face
Pick.PickFaceFromId "Wire_material:solid2", "2"
WCS.AlignWCSWithSelected "Face"
drawBrick_lead(cst_wire_r,wireLength)
'@ pick face
'Pick.PickFaceFromId "Wire_material:solid2", "57"
WCS.Restore("End_of_wires2")
drawBrick_lead(cst_wire_r,wireLength)

ElseIf cst_simp = 1 Then
	If cst_phases_N=3 Then
Pick.PickFaceFromId "Wire_material:solid4", "5"
drawBrick_lead(cst_wire_r,wireLength)
Pick.PickFaceFromId "Wire_material:solid4", "2"
drawBrick_lead(cst_wire_r,wireLength)
End If

Pick.PickFaceFromId "Wire_material:solid2", "5"
drawBrick_lead(cst_wire_r,wireLength)
Pick.PickFaceFromId "Wire_material:solid2", "2"
drawBrick_lead(cst_wire_r,wireLength)
Pick.PickFaceFromId "Wire_material:solid3", "2"
drawBrick_lead(cst_wire_r,wireLength)
Pick.PickFaceFromId "Wire_material:solid3", "5"
drawBrick_lead(cst_wire_r,wireLength)
End If

End Sub

Sub ports()
	Dim kk As Integer


Do While True
kk = kk+1
Debug.Print(kk)
On Error GoTo Handler:
Pick.PickEdgeFromId "Wire_material:Lead"+CStr(kk), "4", "4"
Pick.PickFaceFromId "ground:GND", "2"
Pick.PickFaceFromId "ground:GND", "2"
drawPort(CStr(kk))

If kk=20 Then
	Exit All

End If
Loop
Handler:
Exit All




If 2=1 Then


'_______ports
Pick.PickEdgeFromId "Wire_material:solid6", "4", "4"
Pick.PickFaceFromId "ground:solid11", "2"
Pick.PickFaceFromId "ground:solid11", "2"
drawPort("1")
Pick.PickEdgeFromId "Wire_material:solid5", "4", "4"
Pick.PickFaceFromId "ground:solid11", "2"
drawPort("2")
Pick.PickEdgeFromId "Wire_material:solid7", "4", "4"
Pick.PickFaceFromId "ground:solid11", "2"
drawPort("6")
Pick.PickEdgeFromId "Wire_material:solid10", "4", "4"
Pick.PickFaceFromId "ground:solid11", "2"
drawPort("3")
Pick.PickEdgeFromId "Wire_material:solid9", "4", "4"
Pick.PickFaceFromId "ground:solid11", "2"
drawPort("4")
Pick.PickEdgeFromId "Wire_material:solid8", "4", "4"
Pick.PickFaceFromId "ground:solid11", "2"
drawPort("5")
End If

End Sub

Sub drawBrick_lead(ByVal radius As Double, ByVal height As Double)
	Dim ii As Integer

Handler:
ii=ii+1

With Brick
On Error GoTo Handler:
     .Reset
     .Name "Lead"+CStr(ii)
     .Component "Wire_material"
     .Material "Wire_material"
     .Xrange -0.5*radius, 0.5*radius
     .Yrange -0.5*radius, 0.5*radius
     .Zrange 0, height
     .Create
End With
End Sub

Sub drawGND(ByVal dist As Double, ByVal gnd_x As Double,ByVal gnd_y As Double,ByVal gnd_z As Double)
WCS.Restore "high"
WCS.MoveWCS("local",0,0,dist)
With Brick
     .Reset
     .Name "GND"
     .Component "ground"
     .Material "PEC"
     .Xrange -0.5*gnd_x, 0.5*gnd_x
     .Yrange -0.5*gnd_y, 0.5*gnd_y
     .Zrange 0, gnd_z
     .Create
End With

End Sub

Global Sub drawPort(ByVal num As String)
'WCS.AlignWCSWithSelected "Face"
With DiscreteFacePort
     .Reset
     .PortNumber num
     .Type "SParameter"
     .Impedance "50.0"
     .VoltageAmplitude "1.0"
     '.Current "1.0"
     .SetP1("True", 0, 0 , 0)
     .SetP2("True", 0, 0 , 0)
     .InvertDirection "True"
     .LocalCoordinates "False"
     .Monitor "True"
     '.Radius "0.01"
     .Create
End With

End Sub

Sub setup()

With Material
     .Reset
     .Name "vitroperm"
     .Folder ""
     .FrqType "all"
     .Type "Normal"
     .MaterialUnit "Frequency", "MHz"
     .MaterialUnit "Geometry", "mm"
     .MaterialUnit "Time", "s"
     .MaterialUnit "Temperature", "Kelvin"
     .Epsilon "1"
     .Mu "1"
     .Sigma "0"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstTanD"
     .EnableUserConstTanDModelOrderEps "False"
     .ConstTanDModelOrderEps "1"
     .SetElParametricConductivity "False"
     .ReferenceCoordSystem "Global"
     .CoordSystemType "Cartesian"
     .SigmaM "0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .TanDMModel "ConstTanD"
     .EnableUserConstTanDModelOrderMu "False"
     .ConstTanDModelOrderMu "1"
     .SetMagParametricConductivity "False"
     .DispModelEps  "None"
     .DispModelMu "None"
     .DispersiveFittingSchemeEps "Nth Order"
     .MaximalOrderNthModelFitEps "10"
     .ErrorLimitNthModelFitEps "0.1"
     .UseOnlyDataInSimFreqRangeNthModelEps "False"
     .DispersiveFittingSchemeMu "Nth Order"
     .MaximalOrderNthModelFitMu "10"
     .ErrorLimitNthModelFitMu "0.1"
     .UseOnlyDataInSimFreqRangeNthModelMu "False"
     .UseGeneralDispersionEps "False"
     .DispersiveFittingFormatMu "Real_Imag"
     .AddDispersionFittingValueMu "0.001123324", "9.859266e+04", "3.284724e+03", "1.0"
     .AddDispersionFittingValueMu "0.001261857", "9.856333e+04", "3.686222e+03", "1.0"
     .AddDispersionFittingValueMu "0.001417474", "9.853401e+04", "4.119235e+03", "1.0"
     .AddDispersionFittingValueMu "0.001592283", "9.850469e+04", "4.605742e+03", "1.0"
     .AddDispersionFittingValueMu "0.00178865", "9.847538e+04", "5.170996e+03", "1.0"
     .AddDispersionFittingValueMu "0.002009233", "9.844609e+04", "5.794983e+03", "1.0"
     .AddDispersionFittingValueMu "0.00225702", "9.841680e+04", "6.511523e+03", "1.0"
     .AddDispersionFittingValueMu "0.002535364", "9.838752e+04", "7.207537e+03", "1.0"
     .AddDispersionFittingValueMu "0.002848036", "9.835825e+04", "8.069500e+03", "1.0"
     .AddDispersionFittingValueMu "0.003199267", "9.806460e+04", "9.014004e+03", "1.0"
     .AddDispersionFittingValueMu "0.003593814", "9.733692e+04", "1.014255e+04", "1.0"
     .AddDispersionFittingValueMu "0.004037017", "9.661464e+04", "1.141240e+04", "1.0"
     .AddDispersionFittingValueMu "0.004534879", "9.589772e+04", "1.266798e+04", "1.0"
     .AddDispersionFittingValueMu "0.005094138", "9.518612e+04", "1.401270e+04", "1.0"
     .AddDispersionFittingValueMu "0.005722368", "9.447980e+04", "1.556342e+04", "1.0"
     .AddDispersionFittingValueMu "0.006428073", "9.377872e+04", "1.726724e+04", "1.0"
     .AddDispersionFittingValueMu "0.007220809", "9.308285e+04", "1.906166e+04", "1.0"
     .AddDispersionFittingValueMu "0.008111308", "9.225559e+04", "2.104228e+04", "1.0"
     .AddDispersionFittingValueMu "0.009111628", "9.057879e+04", "2.322807e+04", "1.0"
     .AddDispersionFittingValueMu "0.01023531", "8.858237e+04", "2.548773e+04", "1.0"
     .AddDispersionFittingValueMu "0.01149757", "8.603968e+04", "2.782008e+04", "1.0"
     .AddDispersionFittingValueMu "0.0129155", "8.376049e+04", "3.012658e+04", "1.0"
     .AddDispersionFittingValueMu "0.01450829", "8.082408e+04", "3.243613e+04", "1.0"
     .AddDispersionFittingValueMu "0.01629751", "7.755886e+04", "3.452036e+04", "1.0"
     .AddDispersionFittingValueMu "0.01830738", "7.379304e+04", "3.633635e+04", "1.0"
     .AddDispersionFittingValueMu "0.02056512", "6.985378e+04", "3.810109e+04", "1.0"
     .AddDispersionFittingValueMu "0.0231013", "6.531101e+04", "3.949537e+04", "1.0"
     .AddDispersionFittingValueMu "0.02595024", "6.069582e+04", "4.032430e+04", "1.0"
     .AddDispersionFittingValueMu "0.02915053", "5.640675e+04", "4.068343e+04", "1.0"
     .AddDispersionFittingValueMu "0.03274549", "5.180326e+04", "4.055744e+04", "1.0"
     .AddDispersionFittingValueMu "0.0367838", "4.738032e+04", "3.988130e+04", "1.0"
     .AddDispersionFittingValueMu "0.04132012", "4.333500e+04", "3.863662e+04", "1.0"
     .AddDispersionFittingValueMu "0.04641589", "3.921153e+04", "3.714356e+04", "1.0"
     .AddDispersionFittingValueMu "0.05214008", "3.542173e+04", "3.518628e+04", "1.0"
     .AddDispersionFittingValueMu "0.05857021", "3.199821e+04", "3.316566e+04", "1.0"
     .AddDispersionFittingValueMu "0.06579332", "2.899113e+04", "3.107292e+04", "1.0"
     .AddDispersionFittingValueMu "0.07390722", "2.660282e+04", "2.899692e+04", "1.0"
     .AddDispersionFittingValueMu "0.08302176", "2.441127e+04", "2.705961e+04", "1.0"
     .AddDispersionFittingValueMu "0.09326033", "2.240210e+04", "2.523378e+04", "1.0"
     .AddDispersionFittingValueMu "0.1047616", "2.075892e+04", "2.349589e+04", "1.0"
     .AddDispersionFittingValueMu "0.1176812", "1.923627e+04", "2.187769e+04", "1.0"
     .AddDispersionFittingValueMu "0.1321941", "1.787867e+04", "2.039817e+04", "1.0"
     .AddDispersionFittingValueMu "0.1484968", "1.673415e+04", "1.907854e+04", "1.0"
     .AddDispersionFittingValueMu "0.1668101", "1.566290e+04", "1.784429e+04", "1.0"
     .AddDispersionFittingValueMu "0.1873817", "1.466022e+04", "1.669076e+04", "1.0"
     .AddDispersionFittingValueMu "0.2104904", "1.372173e+04", "1.565948e+04", "1.0"
     .AddDispersionFittingValueMu "0.2364489", "1.281936e+04", "1.469192e+04", "1.0"
     .AddDispersionFittingValueMu "0.2656088", "1.190483e+04", "1.378414e+04", "1.0"
     .AddDispersionFittingValueMu "0.2983647", "1.105554e+04", "1.297067e+04", "1.0"
     .AddDispersionFittingValueMu "0.3351603", "1.026684e+04", "1.221371e+04", "1.0"
     .AddDispersionFittingValueMu "0.3764936", "9.534406e+03", "1.150093e+04", "1.0"
     .AddDispersionFittingValueMu "0.4229243", "8.854222e+03", "1.082974e+04", "1.0"
     .AddDispersionFittingValueMu "0.475081", "8.222563e+03", "1.019773e+04", "1.0"
     .AddDispersionFittingValueMu "0.5336699", "7.591936e+03", "9.597591e+03", "1.0"
     .AddDispersionFittingValueMu "0.5994843", "7.005422e+03", "9.017975e+03", "1.0"
     .AddDispersionFittingValueMu "0.6734151", "6.464219e+03", "8.473362e+03", "1.0"
     .AddDispersionFittingValueMu "0.7564633", "5.964826e+03", "7.961640e+03", "1.0"
     .AddDispersionFittingValueMu "0.8497534", "5.504014e+03", "7.480821e+03", "1.0"
     .AddDispersionFittingValueMu "0.9545485", "5.065912e+03", "7.029040e+03", "1.0"
     .AddDispersionFittingValueMu "1.072267", "4.648638e+03", "6.590421e+03", "1.0"
     .AddDispersionFittingValueMu "1.204504", "4.265735e+03", "6.174705e+03", "1.0"
     .AddDispersionFittingValueMu "1.353048", "3.914371e+03", "5.785211e+03", "1.0"
     .AddDispersionFittingValueMu "1.519911", "3.591949e+03", "5.420287e+03", "1.0"
     .AddDispersionFittingValueMu "1.707353", "3.295047e+03", "5.078382e+03", "1.0"
     .AddDispersionFittingValueMu "1.91791", "2.998154e+03", "4.749597e+03", "1.0"
     .AddDispersionFittingValueMu "2.154435", "2.728011e+03", "4.431981e+03", "1.0"
     .AddDispersionFittingValueMu "2.420128", "2.482209e+03", "4.135606e+03", "1.0"
     .AddDispersionFittingValueMu "2.718588", "2.258554e+03", "3.859049e+03", "1.0"
     .AddDispersionFittingValueMu "3.053856", "2.055052e+03", "3.600987e+03", "1.0"
     .AddDispersionFittingValueMu "3.430469", "1.869886e+03", "3.360181e+03", "1.0"
     .AddDispersionFittingValueMu "3.853529", "1.701403e+03", "3.133297e+03", "1.0"
     .AddDispersionFittingValueMu "4.328761", "1.544661e+03", "2.902665e+03", "1.0"
     .AddDispersionFittingValueMu "4.862602", "1.393798e+03", "2.689009e+03", "1.0"
     .AddDispersionFittingValueMu "5.462277", "1.257668e+03", "2.491079e+03", "1.0"
     .AddDispersionFittingValueMu "6.135907", "1.134835e+03", "2.307719e+03", "1.0"
     .AddDispersionFittingValueMu "6.892612", "1.023998e+03", "2.137855e+03", "1.0"
     .AddDispersionFittingValueMu "7.742637", "9.239861e+02", "1.980494e+03", "1.0"
     .AddDispersionFittingValueMu "8.69749", "8.322074e+02", "1.834716e+03", "1.0"
     .AddDispersionFittingValueMu "9.7701", "7.469682e+02", "1.699668e+03", "1.0"
     .AddDispersionFittingValueMu "10.97499", "6.704598e+02", "1.566485e+03", "1.0"
     .AddDispersionFittingValueMu "12.32847", "6.017877e+02", "1.442666e+03", "1.0"
     .AddDispersionFittingValueMu "13.84886", "5.401494e+02", "1.328633e+03", "1.0"
     .AddDispersionFittingValueMu "15.55676", "4.848244e+02", "1.223614e+03", "1.0"
     .AddDispersionFittingValueMu "17.47528", "4.335097e+02", "1.126896e+03", "1.0"
     .AddDispersionFittingValueMu "19.63041", "3.862294e+02", "1.033178e+03", "1.0"
     .AddDispersionFittingValueMu "22.05131", "3.441056e+02", "9.465589e+02", "1.0"
     .AddDispersionFittingValueMu "24.77076", "3.065760e+02", "8.672014e+02", "1.0"
     .AddDispersionFittingValueMu "27.82559", "2.731396e+02", "7.944971e+02", "1.0"
     .AddDispersionFittingValueMu "31.25716", "2.433499e+02", "7.278882e+02", "1.0"
     .AddDispersionFittingValueMu "35.11192", "2.168049e+02", "6.668636e+02", "1.0"
     .AddDispersionFittingValueMu "39.44206", "1.929075e+02", "6.076024e+02", "1.0"
     .AddDispersionFittingValueMu "44.30621", "1.716442e+02", "5.535972e+02", "1.0"
     .AddDispersionFittingValueMu "49.77024", "1.527247e+02", "5.043921e+02", "1.0"
     .AddDispersionFittingValueMu "55.9081", "1.358906e+02", "4.595605e+02", "1.0"
     .AddDispersionFittingValueMu "62.80291", "1.209120e+02", "4.187136e+02", "1.0"
     .UseGeneralDispersionMu "True"
     .NonlinearMeasurementError "1e-1"
     .NLAnisotropy "False"
     .NLAStackingFactor "1"
     .NLADirectionX "1"
     .NLADirectionY "0"
     .NLADirectionZ "0"
     .Rho "0"
     .ThermalType "Normal"
     .ThermalConductivity "0"
     .HeatCapacity "0"
     .DynamicViscosity "0"
     .Emissivity "0"
     .MetabolicRate "0"
     .BloodFlow "0"
     .VoxelConvection "0"
     .MechanicsType "Unused"
     .Colour "1", "0.501961", "0.501961"
     .Wireframe "False"
     .Reflection "False"
     .Allowoutline "True"
     .Transparentoutline "False"
     .Transparency "0"
     .Create
End With

'@ change material: Ferrite:solid1 to: vitroperm

Solid.ChangeMaterial "Ferrite:solid1", "vitroperm"

'@ change solver type

ChangeSolverType "HF Time Domain"

'@ define tlm solver excitation modes

With TlmSolver
     .ResetExcitationModes
     .SParameterPortExcitation "False"
     .SimultaneousExcitation "True"
     .SetSimultaneousExcitAutoLabel "True"
     .SetSimultaneousExcitationLabel "1[1.0,0.0]+3[1.0,0.0]+5[1.0,0.0]"
     .SetSimultaneousExcitationOffset "Timeshift"
     .PhaseRefFrequency "50.05"
     .ExcitationSelectionShowAdditionalSettings "False"
     .ExcitationPortMode "1", "1", "1.0", "0.0", "default", "True"
     .ExcitationPortMode "2", "1", "1.0", "0.0", "default", "False"
     .ExcitationPortMode "3", "1", "1.0", "0.0", "default", "True"
     .ExcitationPortMode "4", "1", "1.0", "0.0", "default", "False"
     .ExcitationPortMode "5", "1", "1.0", "0.0", "default", "True"
     .ExcitationPortMode "6", "1", "1.0", "0.0", "default", "False"
End With

'@ define time domain solver parameters

Mesh.SetCreator "High Frequency"
With Solver
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "Selected"
     .StimulationMode "All"
     .SteadyStateLimit "-40"
     .MeshAdaption "False"
     .AutoNormImpedance "False"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With


End Sub
