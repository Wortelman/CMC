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
' msgBox "start"
' ResetAll
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
Dim scst_h_gnd As String, cst_h_gnd As Double, portOption As Boolean
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
    portOption      = Evaluate("portOption")
' msgBox CStr(cst_phases_N)
' BeginHide
' assign "cst_phases_N"
' EndHide

On Error Resume Next
 Curve.DeleteCurve "core_curve"
 Curve.DeleteCurve "wire_crosssection"
 Curve.DeleteCurve "torrus_curve"
On Error GoTo 0


'@ define boundaries
With Boundary
     .Xmin "open"
     .Xmax "open"
     .Ymin "open"
     .Ymax "open"
     .Zmin "open"
     .Zmax "electric"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
End With

' eh:
' creationHandling()
'-----------------------------------------------------------------------------------------------------------------------------
'Core drawing
' on Error goto eh
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
	ElseIf cst_phases_N = 3 Then
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

drawlead(0,cst_h_gnd,cst_wire_r,cst_phases_N)

Dim gnd_x As Double
Dim gnd_y As Double
Dim gnd_z As Double
gnd_x=150
gnd_y=150
gnd_z=1

drawGND(cst_h_gnd,gnd_x,gnd_y,gnd_z)

ports(portOption)
setup()



End Sub

Sub drawlead(ByVal cst_simp As Boolean,ByVal cst_h_gnd As Double,ByVal cst_wire_r As Double, ByVal cst_phases_N as Double)
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
'Pick.PickFaceFromId "Wire_material:solid2", "57"
WCS.Restore("End_of_wires2")
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
' WCS.Restore("End_of_wires2")
' drawBrick_lead(cst_wire_r,wireLength)

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

Sub ports(ByVal opt As Boolean)
Dim kk As Integer
if  opt then
'draw 4 ports'
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
' msgBox "error selecting edges 4port"
Exit sub
Else
'draw 2 ports'
Do While True
kk = kk+2
On Error GoTo Handler1:
Pick.PickEdgeFromId "Wire_material:Lead"+CStr(kk-1), "12", "2"
' Pick.PickEdgeFromId "Wire_material:Lead"+CStr(kk-1), "12", "2"
Pick.PickEdgeFromId "Wire_material:Lead"+CStr(kk), "9", "4"
Pick.PickEdgeFromId "Wire_material:Lead"+CStr(kk), "9", "4"
drawPort(CStr(kk/2))
If kk=20 Then
  Exit All
End If
Loop
Handler1:
' msgBox "error selecting edges 2port"
Exit sub

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
Handler2:
ChangeSolverType "HF Time Domain"

End Sub

' Not yet implemented subroutine for clearing Components before creation.
Sub creationHandling ()
dim sname as String
dim parts as Variant
msgBox "got into the creationHandling"
SelectTreeItem("Components\Wire_material\Lead3")
sname = GetSelectedTreeItem
deleteShit(sname)

' Solid.Delete("Wire_material:Lead1")
exit all
' msgBox "got into the creationHandling"
' SelectTreeItem("Components\Ferrite\solid1")
' sname = GetSelectedTreeItem
' msgBox sname
' parts= Split(sname,"\")
' msgBox parts(Ubound(parts))
' Component.Delete(parts(Ubound(parts)))

' parts = splitting(sname,"\")
' msgBox parts(UBound(parts))

' msgBox parts
' msgBox  sname(3)
' Component.Delete(sname(3))
' msgBox "should have deleted something"
' do while sname <> ""
' sname = GetNextSelectedTreeItem 

' Loop
exit all


SelectTreeItem("Components\ground")

SelectTreeItem("Components\Wire_material")

  
End Sub
' Not yet implemented subroutine for clearing Components before creation.
Sub deleteShit(ByVal x as String)
dim stx as String
dim part as Variant
msgBox "got into the deleteShit"
stx = GetSelectedTreeItem
part= Split(stx,"\")
Solid.Delete(part(Ubound(part)-1) +":"+part(Ubound(part)))

do while True
on Error GoTo bla
stx =GetNextSelectedTreeItem 
SelectTreeItem(stx)
part= Split(stx,"\")
Solid.Delete(part(Ubound(part)-1) +":"+part(Ubound(part)))
loop
  
  bla:
  msgBox "error1"
  exit all

End Sub