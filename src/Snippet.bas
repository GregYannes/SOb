' ##########
' ## SOBs ##
' ##########

' ################
' ## SOBs • API ##
' ################

' ###########################
' ## SOBs • API • Creation ##
' ###########################

' Construct a simulated object.
Private Function New_Obj(ByVal cls As String) As Collection
	Obj_Initialize New_Obj, cls
End Function


' Initialize a simulated object.
Private Sub Obj_Initialize(ByRef obj As Collection, _
	ByVal cls As String _
)
	If obj Is Nothing Then
		Set obj = New Collection
	End If
	
	If Not Obj_HasClass(obj) Then
		Obj_Class(obj) = cls
	End If
End Sub



' #########################
' ## SOBs • API • Typing ##
' #########################

' Test for a simulated object.
Private Function IsObj(ByRef x As Variant, _
	Optional ByVal cls As String = VBA.vbNullString _
) As Boolean
' 	Optional ByVal flds() As Long
	
	' Check if the underlying (Collection) structure is correct...
	IsObj = VBA.IsObject(x)
	If IsObj Then
		IsObj = (TypeOf x Is Collection)
	End If
	' IsObj = IsCollection(x)
	
	' ...and that it is marked with a simulated class.
	If IsObj Then
		Dim obj As Object: Set obj = x
		IsObj = Obj_HasClass(obj)
	End If
	
	' Optionally check if the class matches expectations.
	If IsObj And cls <> VBA.VbNullString Then
		IsObj = (Obj_Class(obj) = cls)
	End If
End Function


' Cast as a simulated object.
Private Function AsObj(ByRef x As Variant, _
	Optional ByVal cls As String = VBA.vbNullString _
) As Collection
' 	Optional ByVal flds() As Long
	
	' Cast the underlying structure (to a Collection)...
	Set AsObj = x  ' = AsCollection(x)
	
	' ...and initialize it.
	Obj_Initialize AsObj
	
	' Optionally update the class.
	If cls <> VBA.vbNullString Then
		Obj_Class(x) = cls
	End If
End Function



' #########################
' ## SOBs • API • Fields ##
' #########################
' Count simulated fields.
Private Property Get Obj_FieldCount(ByRef obj As Collection) As Long
' 	Optional ByVal cls As String = VBA.vbNullString
	
	Obj_FieldCount = obj.Count
	
	' Omit the class item from the count of field items.
	If Obj_HasClass(obj) Then
		Obj_FieldCount = Obj_FieldCount - 1
	End If
	
	' Enforce a nonnegative count.
	If Obj_FieldCount < 0 Then
		Obj_FieldCount = 0
	End If
	' Obj_FieldCount = Excel.Application.WorksheetFunction.Max(0, Obj_FieldCount)
End Property


' Test for a simulated field.
Private Function Obj_HasField(ByRef obj As Collection, _
	ByVal fld As Long _
) As Boolean
	Dim key As String: Obj_FieldKey key, fld  ' obj := obj
	Obj_HasField = Clx_Has(obj, key)
End Function
' Get a simulated field.
Private Property Get Obj_Field(ByRef obj As Collection, _
	ByVal fld As Long _
) As Variant
	Dim key As String: Obj_FieldKey key, fld  ' obj := obj
	If Clx_Has(obj, key) Then
		Assign Obj_Field, obj.Item(key)
	End If
End Property


' Set a simulated scalar field...
Private Property Let Obj_Field(ByRef obj As Collection, _
	ByVal fld As Long, _
	ByVal val As Variant _
)
	Dim key As String: Obj_FieldKey key, fld  ' obj := obj
	Clx_Set obj, key, val
End Property
' ...and a simulated scalar field.
Private Property Set Obj_Field(ByRef obj As Collection, _
	ByVal fld As Long, _
	ByRef val As Variant _
)
	Dim key As String: Obj_FieldKey key, fld  ' obj := obj
	Clx_Set obj, key, val
End Property



' ################################
' ## SOBs • API • Visualization ##
' ################################

' Format a simulated object for printing.
Private Function Obj_Format( _
	Optional ByVal cls As String = VBA.vbNullString, _
	Optional ByVal sim As Boolean = False, _
	Optional ByVal ptr As Boolean = False, _
	Optional ByVal sum As String = VBA.vbNullString, _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal ind As String = VBA.vbNullString _
) As String
	' ...
End Function


' .
Private Function Obj_FormatStrInfo( _
	Optional ByVal cls As String = VBA.vbNullString, _
	Optional ByVal ptr As String = VBA.vbNullString, _
	Optional ByVal sum As String = VBA.vbNullString, _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal ind As String = VBA.vbNullString _
) As String
' 	Optional ByVal sim As Boolean = False
' 	Optional ByVal bfr As Boolean = False
	
	Const FMT_IND As String = VBA.vbTab
	Const OBJ_OPEN As String = "<"
	Const OBJ_CLOSE As String = ">"
	Const DTL_SEP As String = ": "
	Const DTL_OPEN As String = "{"
	Const DTL_CLOSE As String = "}"
	Const SUM_SEP As String = ""
	Const SUM_OPEN As String = "["
	Const SUM_CLOSE As String = "]"
	Const PTR_SEP As String = " @ "
	Const PTR_OPEN As String = ""
	Const PTR_CLOSE As String = ""
	
	
	' Assemble the format.
	Dim fmt As String: fmt = ""
	
	
	' Format details across multiple lines...
	If dtl <> VBA.vbNullString Then
		dtl = Application.WorksheetFunction.Clean(dtl)
		
		dtl = DTL_OPEN & VBA.vbNewLine & _
			Text_Indent(dtl, ind := FMT_IND, bfr := True) & VBA.vbNewLine & _
		DTL_CLOSE
		
	' ...or alternatively a summary on a single line...
	ElseIf sum <> VBA.vbNullString Then
		sum = Application.WorksheetFunction.Clean(sum)
		
		' If Text_Contains(sum, VBA.vbNewLine) Then
		' 	sum = VBA.vbNewLine & Text_Indent(sum, ind := FMT_IND, bfr := True) & VBA.vbNewLine
		' End If
		
		sum = SUM_OPEN & sum & SUM_CLOSE
		
	' ...or alternatively a pointer.
	ElseIf ptr <> VBA.vbNullString Then
		ptr = PTR_OPEN & ptr & PTR_CLOSE
	End If
	
	
	' Display only the details in the absence of a class...
	' {
	' 	...
	' 	...
	' }
	If cls = VBA.vbNullString Then
		fmt = dtl
		
	' ...and otherwise enrich the class info:
	ElseIf
		' Name the class itself: <Obj>
		fmt = cls
		
		
		' Append any details:
		' <Obj: {
		' 	...
		' 	...
		' }>
		If dtl <> VBA.vbNullString Then
			fmt = fmt & DTL_SEP & dtl
			
		' Alternatively append a summary: <Obj[...]>
		ElseIf sum <> VBA.vbNullString Then
			fmt = fmt & SUM_SEP & sum
			
		' Alternatively append a pointer: <Obj @ 1234567890>
		ElseIf ptr <> VBA.vbNullString Then
			fmt = fmt & PTR_SEP & ptr
		End If
		
		fmt = OBJ_OPEN & fmt & OBJ_CLOSE
	End If
	
	
	' Propagate any upstream indentation.
	If ind <> VBA.vbNullString Then
		fmt = Text_Indent(fmt, ind := ind, bfr := False)  ' bfr := bfr
	End If
End Function


' Format an array of (simulated) fields for (detailed) printing:
' ".FieldA = True
'  .FieldB = 1
'  .FieldC = 'Yes'
'         ...
'  .FieldZ = <Obj>"
Private Function Obj_FormatFields(ParamArray fields() As Variant) As String
	Const FLD_SEP As String = VBA.vbNewLine
	Const FLD_ARGS As Integer = 2
	
	Dim up As Long: up = UBound(fields, 1)
	Dim low As Long: low = LBound(fields, 1)
	Dim lng As Long: lng = up - low + 1
	Dim n As Long: n = VBA.Int(lng / FLD_ARGS)
	up = n * FLD_ARGS - 1
	
	' Short-circuit for insufficient fields.
	Obj_FormatFields = ""
	If n < 1 Then
		Exit Function
	End If
	
	' Render the first field...
	Dim i As Long: i = low
	Obj_FormatFields = Obj_FormatField(fields(i), fields(i + 1))
	i = i + FLD_ARGS
	
	' ...and append any others.
	For i = i To up Step FLD_ARGS
		Obj_FormatFields = Obj_FormatFields & FLD_SEP & Obj_FormatField(fields(i), fields(i + 1))
	Next i
End Function


' Format a field as an expression for (detailed) printing: ".name = val"
Private Function Obj_FormatField( _
	ByVal name As String, _
	ByVal val As String _
) As String
	Const ASN_OP As String = "="
	Const ASN_SEP As String = " "
	
	' Clean the freeform value for printing.
	val = Application.WorksheetFunction.Clean(val)
	
	' Assemble the format.
	Obj_FormatField = Obj_FormatFieldName(name) & ASN_SEP & ASN_OP & ASN_SEP & val
End Function


' Format the name of a field for printing: ".name"
Private Function Obj_FormatFieldName(ByVal name As String) As String
	Const OBJ_SEP As String = "."
	
	' Clean the name for printing...
	name = Application.WorksheetFunction.Clean(name)
	
	' ...and trim any whitespace.
	name = VBA.Trim(name)
	
	' Assemble the format.
	Obj_FormatFieldName = OBJ_SEP & name
End Function



' ####################
' ## SOBs • Helpers ##
' ####################

' Test for a simulated class.
Private Function Obj_HasClass(ByRef obj As Collection) As Boolean
	Dim key As String: Obj_ClassKey key
	Obj_HasClass = Clx_Has(obj, key)
End Function


' Get the class of a simulated object.
Private Property Get Obj_Class(ByRef obj As Collection) As String
	Dim key As String: Obj_ClassKey key
	If Clx_Has(obj, key) Then
		Obj_Class = obj.Item(key)
	End If
End Property


' Set the class of a simulated object.
Private Property Let Obj_Class(ByRef obj As Collection, _
	ByVal cls As String _
)
	Dim key As String: Obj_ClassKey key
	Clx_Set obj, key, cls
End Property


' Securely obtain the key for a simulated field: "*.Field_i.xxx"
Private Sub Obj_FieldKey(ByRef var As String, _
	ByVal fld As Long _
)
' 	ByRef obj As Collection
' 	ByVal cls As String
	
	Const DEF_CLS As String = VBA.vbNullString
	Const FLD_PFX As String = "Field_"
	Const KEY_SEP As String = "."
	
	Dim cls As String, secret As String, key As String
	' If Obj_HasClass(obj) Then
	' 	cls = Obj_Class(obj)
	' Else
		cls = DEF_CLS
	' End If
	
	Obj_Secret secret
	fld = fld + 1
	
	key = FLD_PFX & VBA.CStr(fld) & KEY_SEP & secret
	If cls <> VBA.vbNullString Then
		key = cls & KEY_SEP & key
	End If
	
	var = key
End Sub


' Securely obtain the key for a simulated class: "Class.xxx"
Private Sub Obj_ClassKey(ByRef var As String)
	Const CLS_PFX As String = "Class"
	Const KEY_SEP As String = "."
	
	Dim key As String, secret As String
	Obj_Secret secret
	key = CLS_PFX & KEY_SEP & secret
	
	var = key
End Sub


' Securely obtain the secret token for keys.
Private Sub Obj_Secret(ByRef var As String)
' 	Optional ByVal refresh As Boolean = False
	
	Const SEC_PFX As String = "x"
	Const REF_SEP As String = ""
	
	Static secret As String, isInit As Boolean
	
	If Not isInit Then  ' Or refresh
		Dim ref1 As New Collection, ref2 As New Collection
		secret = SEC_PFX & VBA.Hex(VBA.ObjPtr(ref1)) & REF_SEP & VBA.Hex(VBA.ObjPtr(ref2))
		
		isInit = True
	End If
	
	var = secret
End Sub



' ######################
' ## SOBs • Utilities ##
' ######################

' Test if a Collection contains an item.
Private Function Clx_Has(ByRef clx As Collection, _
	ByVal index As Variant _
) As Boolean
	Const POS_ERR_NUMBER As Long = 9  ' Subscript out of range.
	Const KEY_ERR_NUMBER As Long = 5  ' Invalid procedure call or argument.
	
	On Error GoTo Fail
	clx.Item index
	
	Clx_Has = True
	Exit Function
	
Fail:
	If VBA.Err.Number = POS_ERR_NUMBER Or VBA.Err.Number = KEY_ERR_NUMBER Then
		Clx_Has = False
	Else
		Err_Raise VBA.Err
	End If
End Function


' Update in a Collection.
Private Sub Clx_Set(ByRef clx As Collection, _
	ByVal key As String, _
	ByRef val As Variant _
)
	If Clx_Has(clx, key) Then
		clx.Remove key
	End If
	' Clx_Remove clx, key
	
	clx.Add val, key := key
End Sub


' Throw an error object.
Private Sub Err_Raise(Optional ByRef e As ErrObject = Nothing)
	If e Is Nothing Then
		Set e = VBA.Err
	End If
	
	VBA.Err.Raise number := e.Number, _
		source := e.Source, _
		description := e.Description, _
		helpFile := e.HelpFile, _
		helpContext := e.HelpContext
End Sub


' Assign a value (scalar or objective) to a variable.
Private Sub Assign( _
	ByRef var As Variant, _
	ByVal val As Variant _
)
	If VBA.IsObject(val) Then
		Set var = val
	Else
		var = val
	End If
End Sub


' ...


' Indent text.
Private Function Text_Indent(ByVal txt As String, _
	Optional ByVal ind As String = VBA.vbTab, _
	Optional ByVal bfr As Boolean = True _
) As String
' 	Optional ByVal old As String = VBA.vbNullString
	
	' Indent the start of every line...
	txt = VBA.Replace(txt, find := VBA.vbNewLine, replace := VBA.vbNewLine & ind)
	
	' ...including (optionally) the beginning.
	If bfr Then
		txt = ind & txt
	End If
	
	' ' Optionally append (rather than prepend) to existing indentation...
	' If old <> VBA.vbNullString Then
	' 	txt = VBA.Replace(txt, find := VBA.vbNewLine & ind & old, replace := VBA.vbNewLine & old & ind)
	' 	
	' 	' ...including (optionally) the beginning.
	' 	If bfr Then
	' 		Dim prePfx As String: prePfx = ind & old
	' 		Dim pfxLen As Long: pfxLen = VBA.Len(prePfx)
	' 		Dim curPfx As String: curPfx = VBA.Left(txt, pfxLen)
	' 		
	' 		If curPfx = prePfx Then
	' 			Dim txtLen As Long: txtLen = VBA.Len(txt)
	' 			Dim sfxLen As Long: sfxLen = txtLen - pfxLen
	' 			Dim sfx As String: sfx = VBA.Right(txt, sfxLen)
	' 			
	' 			txt = old & ind & sfx
	' 		End If
	' 	End If
	' End If
	
	Text_Indent = txt
End Function


' Test if text contains a substring.
Private Function Text_Contains(ByVal txt As String, _
	ByVal sbs As String _
) As Boolean
	Const IDX_NONE As Integer = 0
	
	Text_Contains = (VBA.InStr(string1 := txt, string2 := sbs) <> IDX_NONE)
End Function
