Attribute VB_Name = "SOb"

Option Explicit

' Hide these developer functions from end users in Excel.
Option Private Module



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
		' TODO: Check if this comparison has any side-effects like (property) assignment.
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
	Assign Obj_Field, Clx_Get(obj, key)
End Property


' Set a simulated scalar field...
Private Property Let Obj_Field(ByRef obj As Collection, _
	ByVal fld As Long, _
	ByVal val As Variant _
)
	Dim key As String: Obj_FieldKey key, fld  ' obj := obj
	Clx_Set obj, key, val
End Property
' ...and a simulated objective field.
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
Private Function Obj_Format(ByRef obj As Collection, _
	Optional ByVal dep As Integer = 1, _
	Optional ByVal pln As Boolean = False, _
	Optional ByVal ptr As Boolean = False, _
	Optional ByVal sum As String = VBA.vbNullString, _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal ind As String = VBA.vbNullString, _
	Optional ByVal orf As Boolean = True _
) As String
	Const DFL_CLS As String = "?"
	
	' Format a simulated object...
	If IsObj(obj) Then
		' Extract the class or default to a placeholder ("?").
		Dim cls As String: cls = Obj_Class(obj)
		If cls = VBA.vbNullString Then
			cls = DFL_CLS
		End If
		
		' Optionally record the pointer reference.
		Dim ptrTxt As String: ptrTxt = VBA.vbNullString
		If ptr Then
			ptrTxt = VBA.CStr(VBA.ObjPtr(obj))
		End If
		
		' Format the components.
		Obj_Format = Obj_FormatInfo( _
			cls := cls, _
			dep := dep, _
			pln := pln, _
			ptr := ptrTxt, _
			sum := sum, _
			dtl := dtl, _
			ind := ind, _
			orf := orf _
		)
		
	' ...or display anything else according to VBA defaults.
	Else
		Obj_Format = VBA.TypeName(obj)
	End If
End Function


' Format (simulated) fields for (detailed) printing:
'   .FieldA = True
'   .FieldB = 1
'   .FieldC = 'Yes'
'          ...
'   .FieldZ = <Obj>
'   
' This is done either manually with elegant defaults...
'   Obj_FormatFields0("FieldA", "True", "FieldB", "1", ...)
Private Function Obj_FormatFields0(ParamArray flds() As Variant) As String
	Obj_FormatFields0 = Obj_FormatFields(flds)
End Function


' ...or programmatically with customization:
'   Dim fields() As Variant: fields = Array("FieldA", "True", "FieldB", "1", ...)
'   Obj_FormatFields(fields, vbNewLine)
Private Function Obj_FormatFields( _
	ByRef flds As Variant, _
	Optional ByVal sep As String = VBA.vbNewLine _
) As String
	Const FLD_ARGS As Integer = 2
	
	Dim low As Long: low = LBound(flds, 1)
	Dim up As Long: up = UBound(flds, 1)
	Dim lng As Long: lng = up - low + 1
	Dim n As Long: n = VBA.Int(lng / FLD_ARGS)
	up = n * FLD_ARGS - 1
	
	' Short-circuit for insufficient fields: ""
	If n < 1 Then
		Obj_FormatFields = VBA.vbNullString
		Exit Function
	End If
	
	' Render the first field...
	Dim i As Long: i = low
	Dim fmt As String: fmt = Obj_FormatField(flds(i), flds(i + 1))
	i = i + FLD_ARGS
	
	' ...and append any others.
	For i = i To up Step FLD_ARGS
		fmt = fmt & sep & Obj_FormatField(flds(i), flds(i + 1))
	Next i
	
	Obj_FormatFields = fmt
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
	Obj_Class = Clx_Get(obj, key)
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


' Format a (simulated) object for shallow or deep printing in plain format...
'   {} or {
'   	...
'   	...
'   }
'   
' ...or in rich format:
'   <Obj> or <Obj @1234567890> or <Obj[...]> or
'   <Obj: {
'   	...
'   	...
'   }>
Private Function Obj_FormatInfo( _
	Optional ByVal cls As String = VBA.vbNullString, _
	Optional ByVal dep As Integer = 1, _
	Optional ByVal pln As Boolean = False, _
	Optional ByVal ptr As String = VBA.vbNullString, _
	Optional ByVal sum As String = VBA.vbNullString, _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal ind As String = VBA.vbTab, _
	Optional ByVal orf As Boolean = True _
) As String
	Const OBJ_OPEN As String = "<"
	Const OBJ_CLOSE As String = ">"
	Const DTL_SEP As String = ": "
	Const SUM_SEP As String = ""
	Const SUM_OPEN As String = "["
	Const SUM_CLOSE As String = "]"
	Const PTR_SEP As String = " "
	Const PTR_OPEN As String = "@"
	Const PTR_CLOSE As String = ""
	
	' Sanitize depth.
	dep = Excel.Application.WorksheetFunction.Max(0, dep)
	
	' Format shallowly when details are absent.
	If dtl = VBA.vbNullString Then
		dep = 0
	End If
	
	' Assemble plain formatting...
	Dim fmt As String
	If pln Then		
		' Format deeply...
		'   {
		'   	...
		'   	...
		'   }
		If dep > 0 Then
			' dtl = Excel.Application.WorksheetFunction.Clean(dtl)
			fmt = Obj_FormatDetails(dtl, ind := ind, orf := orf)
			
		' ...or shallowly: {}
		Else
			fmt = Obj_FormatDetails()
		End If
		
	' ...or rich formatting.
	Else
		' Short circuit for missing cls: ""
		If cls = VBA.vbNullString Then
			Obj_FormatInfo = VBA.vbNullString
			Exit Function
		End If
		
		' Clean the class name.
		cls = Excel.Application.WorksheetFunction.Clean(cls)
		cls = VBA.Trim(cls)
		
		fmt = cls
		
		' Format deeply with details...
		'   <Obj: {
		'   	...
		'   	...
		'   }>
		If dep > 0 Then
			' dtl = Excel.Application.WorksheetFunction.Clean(dtl)
			fmt = cls & DTL_SEP & Obj_FormatDetails(dtl, ind := ind, orf := orf)
			
		' ...or shallowly...
		Else
			' ...with maybe a summary: <Obj[...]>
			If sum <> VBA.vbNullString Then
				sum = Excel.Application.WorksheetFunction.Clean(sum)
				fmt = cls & SUM_SEP & SUM_OPEN & sum & SUM_CLOSE
				
			' ...or maybe a pointer: <Obj @1234567890>
			ElseIf ptr <> VBA.vbNullString Then
				ptr = Excel.Application.WorksheetFunction.Clean(ptr)
				ptr = VBA.Trim(ptr)
				fmt = cls & PTR_SEP & PTR_OPEN & ptr & PTR_CLOSE
				
			' ...or with only the name: <Obj>
			End If
		End If
		
		' Wrap result.
		fmt = OBJ_OPEN & fmt & OBJ_CLOSE
	End If
	
	Obj_FormatInfo = fmt
End Function


' Format details for printing.
'   {} or {...} or {
'   	...
'   	...
'   }
Private Function Obj_FormatDetails( _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal ind As String = VBA.vbTab, _
	Optional ByVal orf As Boolean = True _
) As String
	Const DTL_OPEN As String = "{"
	Const DTL_CLOSE As String = "}"
	
	' Indent details on separate lines: not for missing details ("{}")...
	Dim brk As Boolean
	If dtl = VBA.vbNullString Then
		brk = False
	Else
		' ...but certainly for multiline details...
		'   {
		'   	...
		'   	...
		'   }
		If Text_Contains(dtl, VBA.vbNewLine) Then
			brk = True
			
		' ...and optionally for orphan lines:
		'   {
		'   	...
		'   }
		ElseIf orf Then
			brk = True
		Else
			brk = False
		End If
	End If
	
	If brk Then
		dtl = VBA.vbNewLine & Text_Indent(dtl, ind := ind, bfr := True) & VBA.vbNewLine
	End If
	
	' Wrap detail in braces.
	dtl = DTL_OPEN & dtl & DTL_CLOSE
	
	Obj_FormatDetails = dtl
End Function


' Format a field as an expression for (detailed) printing: .name = val
Private Function Obj_FormatField( _
	ByVal name As String, _
	ByVal val As String _
) As String
	Const OBJ_SEP As String = "."
	Const ASN_OP As String = "="
	Const ASN_SEP As String = " "
	
	' Clean the name for printing...
	name = Excel.Application.WorksheetFunction.Clean(name)
	name = VBA.Trim(name)
	
	' ' ...along with the freeform value for printing.
	' val = Excel.Application.WorksheetFunction.Clean(val)
	
	' Assemble the format.
	Obj_FormatField = OBJ_SEP & name & ASN_SEP & ASN_OP & ASN_SEP & val
End Function



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


' Get an item (safely) from a Collection.
Private Function Clx_Get(ByRef clx As Collection, _
	ByVal index As Variant _
) As Variant
	If Clx_Has(clx, index) Then
		Assign Clx_Get, clx.Item(index)
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
