'function to check if input is in specified range return either True or False 
Private Function InRange(Range1 As Range, Range2 As Range) As Boolean
    ' returns True if Range1 is within Range2
    InRange = Not (Application.Intersect(Range1, Range2) Is Nothing)
End Function

'function to check if copied element is text, return True or False
Function IsTextInClipBoard() As Boolean
    Dim sClipText As String
    
    On Error Resume Next
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        sClipText = .GetText
        IsTextInClipBoard = sClipText <> vbNullString
    End With

End Function


'shortcuts enable  Call this function via button on the ribbon
Sub PPARenableshorcut()

	Application.OnKey "{F6}", "Insertphoto"

	MsgBox ("Macros enabled:" & vbCrLf & vbCrLf & "F6 - Paste and resize photo")

End Sub
'shortcuts enable

Sub Insertphoto()

    Dim myShape As Excel.shape
    Dim width As Single
    Dim height As Single
    Dim thePicture As shape
    Dim wrongcell
    
    wrongcell = 0  'reset variable
	
	If IsTextInClipBoard = False Then 'first check if copied element is picture (true = text, false = not text)
			If InRange(ActiveCell, Range("G2:H10")) Then 'check if sected cell is in range G2:H10
				wrongcell = 1 'selection check
				Range("G3").Select  'select Top-left corner of selection
				ActiveSheet.Paste	'paste picture
					'call function to resize picture to the size of the provided width and height 
					Call ResizePicture(Range("G2:H10"), 138, 122, False) 'range, width, height, isExtra (for additional function - True/False)
			End If

			'same as above but with True for additional function (in this case set the picture to monochromatic)
			If InRange(ActiveCell, Range("J1:N50")) Then 
				wrongcell = 1 
				Range("J1").Select
				ActiveSheet.Paste
				Call ResizePicture(Range("J1:N50"), 800, 400, True) 
				
				'extra - drawing RED circle on top of pasted picture
				ActiveSheet.Shapes.AddShape(msoShapeOval, 725.2940944882, 154.4117322835, _
				30, 30).Select
			With Selection.ShapeRange.Fill
				.Visible = msoTrue
				.ForeColor.RGB = RGB(255, 0, 0)
				.Transparency = 0
				.Solid
			End With
			With Selection.ShapeRange.Line
				.Visible = msoTrue
				.ForeColor.RGB = RGB(255, 0, 0)
				.Transparency = 0
			End With
			End If
		 
			'if selected cells is not present in any of above ranges return message to the user
		If wrongcell = 0 Then
			MsgBox ("Selected cell not present in suported range")
		End If
	Else
		MsgBox ("Copied element is not a picture")
	End If
End Sub


'function to resize pictures (and more)
Sub ResizePicture(rangepaste As Range, towidth As Integer, toheight As Integer, isextra As Boolean)
    Dim shape As shape
    Dim scalefactor
    Dim scalefactorsecond

For Each shape In ActiveSheet.Shapes  'find picture in defined range
    If Not Application.Intersect(shape.TopLeftCell, rangepaste) Is Nothing Then 'if found Do below
        currentwidth = shape.width         'get current width of the picture
        currentheight = shape.height       'get current height of the picture
        
        If isextra = False Then				'check if additional function is needed
        
			'do if no additional function is needed
            If currentwidth >= currentheight Then  'check height and width which one is bigger, do if width is bigger
                scalefactor = towidth / currentwidth  'set new scale by dividing expected width by current width
                shape.ScaleWidth scalefactor, msoFalse, msoScaleFromTopLeft
				'scaling as per scalefactor msoFalse=scaling current size of the pizcutre, msoScalefromTopLeft=which border will stay in place
            Else
                scalefactor = toheight / currentheight 'as above but for height
                shape.ScaleHeight scalefactor, msoFalse, msoScaleFromTopLeft 'as above but for height
            End If
            
        Else 'if extra function is needed do below
            shape.Fill.PictureEffects.Insert(msoEffectSaturation).EffectParameters(1).Value = 0  'set saturation of picture to 0, AKA monochromatic
            scalefactor = toheight / currentheight 'set scale (as above)
            shape.ScaleHeight scalefactor, msoFalse, msoScaleFromTopLeft 'scale by height, in my case I needed only by height, if by width is also required you might need to rewrite this or add by width baseing on this
            
        End If 'end of isextra
        
        If shape.width <= towidth Then  'one more time check if resized picture fits provided range
        
            If shape.height <= toheight Then
            
            Else 'if it doesnt fit in height, resize it again
                scalefactorsecond = toheight / shape.height   
                shape.ScaleHeight scalefactorsecond, msoFalse, msoScaleFromTopLeft
            End If
            
        Else 'if it doesnt fit in width, resize it again
            scalefactorsecond = towidth / shape.width  
            shape.ScaleWidth scalefactorsecond, msoFalse, msoScaleFromTopLeft
        End If
        Exit For 'Required, otherwise it will keep checking your range for picture and will throw error 1004
    End If
Next shape
End Sub
