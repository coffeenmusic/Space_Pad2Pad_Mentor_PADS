' Description

' Useful commands
' ProcessParameter(45 , "w")
' ActiveDocument.SelectObjects(Type:=_, value:=_, Select:=_)
' SendKeys "{Home}+{End}"

' Public Const spacing = 40

Dim axis As String
Dim selection(100) As String
Dim spacing As Single

Public Function Get_Selected_Pins () As Integer
	cnt = 0
	For Each p In ActiveDocument.Pins
		If p.selected Then
			selection(cnt) = p.Name
			cnt = cnt + 1
		End If
	Next p
	
	Get_Selected_Pins = cnt
End Function

Public Function Move_Selected_Cmp (axis, value)
	For Each cmp In ActiveDocument.Components
		If cmp.selected Then
			If axis = "x" Then
				cmp.Move(value, cmp.PositionY, 0)
			Else
				cmp.Move(cmp.PositionX, value, 0)
			End If			
		End If
	Next cmp
End Function

Public Function Get_Pad_Delta(axis, p)
	l = 0
	w = 0
	r = 0
	For Each ps In p.PadStackLayers
		If ps.Number = ppcbPadStackLayerTypeMounted Then
			l = ps.Pad.Length
			w = ps.Pad.width
			r = Abs(ps.Pad.Orientation - p.Component.Orientation)
		End If
	Next ps
	
	If axis = "x" Then
		If r = 90 Or r = 270 Then
			Get_Pad_Delta = w
		Else
			Get_Pad_Delta = l
		End If
	Else
		If r = 90 Or r = 270 Then
			Get_Pad_Delta = l
		Else
			Get_Pad_Delta = w
		End If
	End If
End Function

Public Function Get_Closest_Pin(axis, p_pin) As String
	smallest_delta = 999999.9
	
	If axis = "x" Then
		pcoord = p_pin.PositionX
	Else
		pcoord = p_pin.PositionY
	End If
	
	For Each cmp In ActiveDocument.Components
		If cmp.selected Then
			For Each p In cmp.Pins
				If axis = "x" Then
					coord = p.PositionX
				Else
					coord = p.PositionY
				End If
				
				delta = Abs(coord - pcoord)
				If delta < smallest_delta Then
					smallest_delta = delta
					closest_pin = p.Name
				End If
			Next p	
		End If
	Next cmp
	
	Get_Closest_Pin = closest_pin
End Function

Public Function Get_Direction(axis, p_pin, p_cmp) As Integer
	direction = 1
	If axis = "x" Then
		coord_cmp = p_cmp.PositionX
		coord_pin = p_pin.PositionX
	Else
		coord_cmp = p_cmp.PositionY
		coord_pin = p_pin.PositionY
	End If
	
	If coord_cmp - coord_pin >= 0 Then
		Get_Direction = 1
	Else
		Get_Direction = -1
	End If
End Function

Public Function Get_New_Cmp_Coord(axis, p_pin, p_cmp, total_delta) As Single
	If axis = "x" Then
		coord = p_pin.PositionX + total_delta
		delta = coord - p_cmp.PositionX
		Get_New_Cmp_Coord = p_cmp.Component.PositionX + delta
	Else
		coord = p_pin.PositionY + total_delta
		delta = coord - p_cmp.PositionY
		Get_New_Cmp_Coord = p_cmp.Component.PositionY + delta
	End If
End Function

Sub Move_Component (axis)
	Dim coord As Single
	
	cnt = Get_Selected_Pins
	If cnt > 1 Then
		MsgBox("Please select only 1 pin")
		GoTo EndScript
	End If
	
	
	Set p_pin = ActiveDocument.Pins(selection(0)) ' Pin Object
	p_pin_delta = Get_Pad_Delta (axis, p_pin)	
	closest_pin = Get_Closest_Pin(axis, p_pin)
	
	Set p_cmp = ActiveDocument.Pins(closest_pin) ' Pin Object
	p_cmp_delta = Get_Pad_Delta (axis, p_cmp)
	
	direction = Get_Direction(axis, p_pin, p_cmp)
	total_delta = direction * (p_pin_delta/2 + p_cmp_delta/2 + spacing)	
	
	' Get new coordinate and move component to it
	new_coord = Get_New_Cmp_Coord(axis, p_pin, p_cmp, total_delta)
	Move_Selected_Cmp (axis, new_coord)
	
	EndScript:
		Exit All
End Sub

Sub Main
	On Error GoTo EndScript
	Begin Dialog UserDialog 400,105,"Move component so Pad to Pad spacing is as defined",.CallbackFunc ' %GRID:10,7,1,1
		TextBox 100,42,270,21,.txtSpacing
		Text 30,44,60,14,"Spacing",.lblSpacing
		CancelButton 220,77,150,21
		OptionGroup .gbAxis
			OptionButton 180,14,40,14,"X",.rbX
			OptionButton 240,14,40,14,"Y",.rbY
		OKButton 40,77,130,21,.btnOK
	End Dialog
	Dim dlg As UserDialog
	DialogOK = Dialog(dlg)
	
	If DialogOK Then
		spacing_str = dlg.txtSpacing
		If spacing_str = "" Then
			MsgBox("Spacing text box must not be empty.")
			GoTo EndScript
		End If
		spacing = CDbl(spacing_str)
	End If

	Move_Component axis
	
	EndScript:
		Exit All
End Sub

Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
		Case 1 ' Dialog box initialization
			axis = "x"
			DlgValue "gbAxis",0
		Case 2 ' Value changing or button pressed
			If ((DlgItem$ <> "btnOK") And (DlgItem$ <> "Cancel")) Then
				CallbackFunc = True ' Prevent button press from closing the dialog box
				If (DlgItem$ = "gbAxis") Then
					Select Case SuppValue%
						Case 0
							axis = "x"
						Case 1
							axis = "y"
					End Select
				End If
			ElseIf DlgItem$ = "Cancel" Then
			 	CallbackFunc = False ' closing the dialog box
			End If
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function

