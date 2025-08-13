Attribute VB_Name = "SOb"

Option Explicit

' Hide these developer functions from end users in Excel.
Option Private Module



' ##############
' ## Metadata ##
' ##############

' Module info.
Public Const MOD_NAME As String = "SOb"

Public Const MOD_VERSION As String = "0.1.0"

Public Const MOD_REPO As String = "https://github.com/GregYannes/SOb"



' #########
' ## API ##
' #########

' ####################
' ## API | Creation ##
' ####################

' Construct a simulated object.
Public Function New_Obj(ByVal class As String) As Collection
	Obj_Initialize New_Obj, class
End Function


' Initialize a simulated object.
Public Sub Obj_Initialize(ByRef obj As Collection, _
	ByVal class As String _
)
	If obj Is Nothing Then
		Set obj = New Collection
	End If
	
	' Update the class if it is missing...
	If Not Obj_HasClass(obj) Then
		Obj_Class(obj) = class
		
	' ...or is blank.
	ElseIf Obj_Class(obj) = VBA.vbNullString Then
		Obj_Class(obj) = class
	End If
End Sub



' ####################
' ## API | Typology ##
' ####################

' The class of a simulated object: the developer may read it...
Public Property Get Obj_Class(ByRef obj As Collection) As String
	Dim key As String: Obj_ClassKey key
	Obj_Class = Clx_Get(obj, key)
End Property


' ...but never write it directly.
Private Property Let Obj_Class(ByRef obj As Collection, _
	ByVal class As String _
)
	class = Excel.Application.WorksheetFunction.Clean(class)
	class = VBA.Trim(class)
	
	Dim key As String: Obj_ClassKey key
	Clx_Set obj, key, class
End Property


' Test for a simulated object.
Public Function IsObj(ByRef x As Variant, _
	Optional ByVal class As String = VBA.vbNullString, _
	Optional ByRef fields As Variant, _
	Optional ByVal strict As Boolean = True _
) As Boolean
	' Check if the underlying (Collection) structure is correct...
	IsObj = VBA.IsObject(x)
	If Not IsObj Then Exit Function
	
	Dim obj As Object: Set obj = x
	IsObj = (TypeOf obj Is Collection)
	If Not IsObj Then Exit Function
	
	' ...and that it is marked with a simulated class.
	IsObj = Obj_HasClass(obj)
	If Not IsObj Then Exit Function
	
	' Optionally check if the class matches expectations.
	If class <> VBA.VbNullString Then
		IsObj = (Obj_Class(obj) = class)
	End If
	If Not IsObj Then Exit Function
	
	' Optionally check for the presence of specific fields.
	If Not VBA.IsMissing(fields) Then
		IsObj = Obj_HasFields(obj, fields := fields)
		
		' Optionally cap the fields at strictly those specified, rather than a superset.
		If strict Then
			Dim nFlds As Long: nFlds = Arr_Length(fields, 1)
			IsObj = (Obj_FieldCount(obj) = nFlds)
		End If
	End If
	If Not IsObj Then Exit Function
End Function


' Cast as a simulated object.
Public Function AsObj(ByRef x As Variant, _
	Optional ByVal class As String = VBA.vbNullString _
) As Collection
	' Cast the underlying structure (to a Collection)...
	Set AsObj = x
	
	' ...and initialize it.
	If class = VBA.vbNullString Then
		Obj_Initialize AsObj, VBA.vbNullString
	Else
		Obj_Initialize AsObj, class
	End If
	
	' Optionally update the class.
	If class <> VBA.vbNullString Then
		Obj_Class(AsObj) = class
	End If
End Function



' ##################
' ## API | Fields ##
' ##################

' Get a simulated field as a Property.
Public Property Get Obj_Field(ByRef obj As Collection, _
	ByVal field As Long _
) As Variant
	Dim key As String: Obj_FieldKey key, field
	Assign Obj_Field, Clx_Get(obj, key)
End Property


' Set a simulated scalar field as a Property...
Public Property Let Obj_Field(ByRef obj As Collection, _
	ByVal field As Long, _
	ByVal val As Variant _
)
	Dim key As String: Obj_FieldKey key, field
	Clx_Set obj, key, val
End Property


' ...and a simulated objective field.
Public Property Set Obj_Field(ByRef obj As Collection, _
	ByVal field As Long, _
	ByRef val As Variant _
)
	Dim key As String: Obj_FieldKey key, field
	Clx_Set obj, key, val
End Property


' Safely get a simulated (Property) field.
Public Sub Obj_Get(ByRef var As Variant, _
	ByRef obj As Collection, _
	ByVal field As Long _
)
	Dim has As Boolean
	Dim key As String: Obj_FieldKey key, field
	Dim val As Variant: Assign val, Clx_Get(obj, key, has := has)
	
	' Store the value when the field is present.
	If has Then
		Assign var, val
	End If
End Sub


' Count simulated fields.
Public Function Obj_FieldCount(ByRef obj As Collection) As Long
	Obj_FieldCount = obj.Count
	
	' Omit the class item from the count of field items.
	If Obj_HasClass(obj) Then
		Obj_FieldCount = Obj_FieldCount - 1
	End If
	
	' Enforce a nonnegative count.
	If Obj_FieldCount < 0 Then
		Obj_FieldCount = 0
	End If
End Function


' Test for a single simulated field.
Public Function Obj_HasField(ByRef obj As Collection, _
	ByVal field As Long _
) As Boolean
	Dim key As String: Obj_FieldKey key, field
	Obj_HasField = Clx_Has(obj, key)
End Function


' Test programmatically for multiple simulated fields...
Private Function Obj_HasFields(ByRef obj As Collection, _
	ByRef fields As Variant _
) As Boolean
	Dim n As Long: n = Arr_Length(fields, 1)
	
	' Short-circuit with TRUE for trivial case.
	If n < 1 Then
		Obj_HasFields = True
		Exit Function
	End If
	
	Dim low As Long: low = LBound(fields, 1)
	Dim up As Long: up = UBound(fields, 1)
	
	' Return FALSE if any fields are nonexistent...
	Dim i As Long
	For i = low To up
		If Not Obj_HasField(obj, fields(i)) Then
			Obj_HasFields = False
			Exit Function
		End If
	Next i
	
	' ...and otherwise return TRUE since they all exist.
	Obj_HasFields = True
End Function


' ...or test manually.
Public Function Obj_HasFields0(ByRef obj As Collection, _
	ParamArray fields() As Variant _
) As Boolean
	Dim f() As Variant: f = fields
	Obj_HasFields0 = Obj_HasFields(obj, fields := f)
End Function



' ######################
' ## API | Validation ##
' ######################

' Checks that simulated fields match the type constraints of their accessors.
Public Sub Obj_Check(ParamArray fields() As Variant)
End Sub


' Catches errors for certain checks and propagates all others.
Public Function Obj_CheckError(Optional ByRef e As ErrObject = Nothing, _
	Optional ByVal type_ As Boolean = True _
) As Boolean
	Const NO_ERR_NUMBER As Integer = 0         ' No error.
	Const TYP_OBJ_ERR_NUMBER As Integer = 13   ' Invalid type.
	Const TYP_SCL_ERR_NUMBER As Integer = 450  ' Wrong number of arguments or invalid property assignment.
	
	' Default to last error.
	If e Is Nothing Then
		Set e = VBA.Err
	End If
	
	' Handle various errors.
	Dim cat As Boolean
	Select Case e.Number
		' Short-circuit with TRUE for no error.
		Case NO_ERR_NUMBER
			Obj_CheckError = True
			Exit Function
			
		' Mark specific errors for catching as desired: namely mismatched data types.
		Case TYP_SCL_ERR_NUMBER, TYP_OBJ_ERR_NUMBER
			cat = type_
			
		' Mark all other errors for propagation.
		Case Else
			cat = False
	End Select
	
	' Return FALSE for errors that should be caught...
	If cat Then
		Obj_CheckError = False
		
	' ...and propagate all others.
	Else
		Err_Raise e
	End If
End Function



' #########################
' ## API | Visualization ##
' #########################

' Print a simulated object with automatic formatting...
Public Function Obj_Print(ByRef obj As Collection, _
	Optional ByVal depth As Integer = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal ptr As Boolean = False, _
	Optional ByVal sum As String = VBA.vbNullString, _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal pvw as Boolean = False, _
	Optional ByVal ind As String = VBA.vbTab, _
	Optional ByVal orf As Boolean = True _
) As String
	Obj_Print = Obj_Format(obj, _
		depth := depth, _
		plain := plain, _
		ptr := ptr, _
		sum := sum, _
		dtl := dtl, _
		pvw := pvw, _
		ind := ind, _
		orf := orf _
	)
	
	Obj_Print0 Obj_Print
End Function


' ...or verbatim.
Public Function Obj_Print0(Optional ByRef fmt As String = VBA.vbNullString) As String
	Obj_Print0 = fmt
	
	Debug.Print Obj_Print0
End Function


' Format a simulated object for printing.
Public Function Obj_Format(ByRef obj As Collection, _
	Optional ByVal depth As Integer = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal ptr As Boolean = False, _
	Optional ByVal sum As String = VBA.vbNullString, _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal pvw as Boolean = False, _
	Optional ByVal ind As String = VBA.vbTab, _
	Optional ByVal orf As Boolean = True _
) As String
	Const DFL_CLS As String = "?"
	
	' Format a simulated object...
	If IsObj(obj) Then
		' Extract any class...
		Dim cls As String
		If Obj_HasClass(obj) Then
			cls = Obj_Class(obj)
		Else
			cls = VBA.vbNullString
		End If
		
		' ...and default to a placeholder for an unknown: "?"
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
			class := cls, _
			depth := depth, _
			plain := plain, _
			ptr := ptrTxt, _
			sum := sum, _
			dtl := dtl, _
			pvw := pvw, _
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
' This is done either programmatically with customization...
'   Dim fields() As Variant: fields = Array("FieldA", "True", "FieldB", "1", ...)
'   Obj_FormatFields(fields, vbNewLine)
Public Function Obj_FormatFields( _
	ByRef fields As Variant, _
	Optional ByVal sep As String = VBA.vbNewLine _
) As String
	Const FLD_ARGS As Integer = 2
	
	Dim lng As Long: lng = Arr_Length(fields, 1)
	Dim n As Long: n = VBA.Int(lng / FLD_ARGS)
	
	' Short-circuit for insufficient fields: ""
	If n < 1 Then
		Obj_FormatFields = VBA.vbNullString
		Exit Function
	End If
	
	Dim low As Long: low = LBound(fields, 1)
	Dim up As Long: up = n * FLD_ARGS - 1
	
	' Render the first field...
	Dim i As Long: i = low
	Dim fmt As String: fmt = Obj_FormatField(fields(i), fields(i + 1))
	i = i + FLD_ARGS
	
	' ...and append any others.
	For i = i To up Step FLD_ARGS
		fmt = fmt & sep & Obj_FormatField(fields(i), fields(i + 1))
	Next i
	
	Obj_FormatFields = fmt
End Function


' ...or manually with elegant defaults.
'   Obj_FormatFields0("FieldA", "True", "FieldB", "1", ...)
Public Function Obj_FormatFields0(ParamArray fields() As Variant) As String
	Dim f() As Variant: f = fields
	Obj_FormatFields0 = Obj_FormatFields(f)
End Function



' #############
' ## Support ##
' #############

' Test for a simulated class.
Private Function Obj_HasClass(ByRef obj As Collection) As Boolean
	Dim key As String: Obj_ClassKey key
	Obj_HasClass = Clx_Has(obj, key)
End Function


' Securely obtain the key for a simulated field: "*.Field_i.xxx"
Private Sub Obj_FieldKey(ByRef var As String, _
	ByVal field As Long _
)
	Const FLD_PFX As String = "Field_"
	Const KEY_SEP As String = "."
	
	Dim sec As String, key As String
	
	Obj_Secret sec
	field = field + 1
	
	key = FLD_PFX & VBA.CStr(field) & KEY_SEP & sec
	
	var = key
End Sub


' Securely obtain the key for a simulated class: "Class.xxx"
Private Sub Obj_ClassKey(ByRef var As String)
	Const CLS_PFX As String = "Class"
	Const KEY_SEP As String = "."
	
	Dim key As String, sec As String
	Obj_Secret sec
	key = CLS_PFX & KEY_SEP & sec
	
	var = key
End Sub


' Securely obtain the secret token for keys.
Private Sub Obj_Secret(ByRef var As String)
	Const SEC_PFX As String = "x"
	Const REF_SEP As String = ""
	
	Static sec As String, isInit As Boolean
	
	If Not isInit Then
		Dim ref1 As New Collection, ref2 As New Collection
		sec = SEC_PFX & VBA.Hex(VBA.ObjPtr(ref1)) & REF_SEP & VBA.Hex(VBA.ObjPtr(ref2))
		
		isInit = True
	End If
	
	var = sec
End Sub


' Format a (simulated) object for shallow or deep printing in plain format...
'   {} or {…} or {
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
	Optional ByVal class As String = VBA.vbNullString, _
	Optional ByVal depth As Integer = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal ptr As String = VBA.vbNullString, _
	Optional ByVal sum As String = VBA.vbNullString, _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal pvw As Boolean = False, _
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
	If depth < 0 Then
		depth = 0
	End If
	
	' Assemble plain formatting...
	Dim fmt As String
	If plain Then		
		' Format deeply...
		'   {
		'   	...
		'   	...
		'   }
		If depth > 0 Then
			fmt = Obj_FormatDetails(dtl, pvw := False, ind := ind, orf := orf)
			
		' ...or shallowly: {…} or {}
		Else
			fmt = Obj_FormatDetails(dtl, pvw := True)
		End If
		
	' ...or rich formatting.
	Else
		' Short circuit for missing class: ""
		If class = VBA.vbNullString Then
			Obj_FormatInfo = VBA.vbNullString
			Exit Function
		End If
		
		' Clean the class name.
		class = Excel.Application.WorksheetFunction.Clean(class)
		class = VBA.Trim(class)
		
		fmt = class
		
		' Format deeply with details...
		'   <Obj: {
		'   	...
		'   	...
		'   }>
		If depth > 0 Then
			fmt = class & DTL_SEP & Obj_FormatDetails(dtl, pvw := False, ind := ind, orf := orf)
			
		' ...or shallowly...
		Else
			' ...with maybe a summary: <Obj[...]>
			If sum <> VBA.vbNullString Then
				sum = Excel.Application.WorksheetFunction.Clean(sum)
				fmt = class & SUM_SEP & SUM_OPEN & sum & SUM_CLOSE
				
			' ...or maybe a preview of the detail: <Obj: {…}>
			ElseIf pvw Then
				fmt = class & DTL_SEP & Obj_FormatDetails(dtl, pvw := True)
				
			' ...or maybe a pointer: <Obj @1234567890>
			ElseIf ptr <> VBA.vbNullString Then
				ptr = Excel.Application.WorksheetFunction.Clean(ptr)
				ptr = VBA.Trim(ptr)
				fmt = class & PTR_SEP & PTR_OPEN & ptr & PTR_CLOSE
				
			' ...or with only the name: <Obj>
			End If
		End If
		
		' Wrap result.
		fmt = OBJ_OPEN & fmt & OBJ_CLOSE
	End If
	
	Obj_FormatInfo = fmt
End Function


' Format details for printing.
'   {} or {…} or {...} or {
'   	...
'   	...
'   }
Private Function Obj_FormatDetails( _
	Optional ByVal dtl As String = VBA.vbNullString, _
	Optional ByVal pvw As Boolean = False, _
	Optional ByVal ind As String = VBA.vbTab, _
	Optional ByVal orf As Boolean = True _
) As String
	Const DTL_OPEN As String = "{"
	Const DTL_CLOSE As String = "}"
	
	' Horizontal ellipsis: "…"
	#If Mac Then
		Const DTL_PVW As Long = 201
	#Else
		Const DTL_PVW As Long = 133
	#End If
	
	' Optionally show only a preview ("{…}") of the details...
	If pvw Then
		If dtl <> VBA.vbNullString Then
			dtl = VBA.Chr(DTL_PVW)
		End If
		
	' ...and otherwise break out details.
	Else
		Dim brk As Boolean
		
		' Not for missing details ("{}")...
		If dtl = VBA.vbNullString Then
			brk = False
			
		' ...but certainly for multiline details...
		'   {
		'   	...
		'   	...
		'   }
		ElseIf Txt_Contains(dtl, VBA.vbNewLine) Then
			brk = True
			
		' ...and optionally for orphan lines...
		'   {
		'   	...
		'   }
		Else
			brk = orf
		End If
		
		' Indent as needed.
		If brk Then
			dtl = VBA.vbNewLine & Txt_Indent(dtl, ind := ind, bfr := True) & VBA.vbNewLine
		End If
	End If
	
	' Wrap details in braces.
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
	
	' Assemble the format.
	Obj_FormatField = OBJ_SEP & name & ASN_SEP & ASN_OP & ASN_SEP & val
End Function



' ###############
' ## Utilities ##
' ###############

' Assign a value (scalar or objective) to a variable.
Public Sub Assign( _
	ByRef var As Variant, _
	ByVal val As Variant _
)
	If VBA.IsObject(val) Then
		Set var = val
	Else
		Let var = val
	End If
End Sub


' Indent text.
Public Function Txt_Indent(ByVal txt As String, _
	Optional ByVal ind As String = VBA.vbTab, _
	Optional ByVal bfr As Boolean = True _
) As String
	' Indent the start of every line...
	txt = VBA.Replace(txt, find := VBA.vbNewLine, replace := VBA.vbNewLine & ind)
	
	' ...including (optionally) the beginning.
	If bfr Then
		txt = ind & txt
	End If
	
	Txt_Indent = txt
End Function


' Test if a Collection contains an item.
Private Function Clx_Has(ByRef clx As Collection, _
	ByVal index As Variant _
) As Boolean
	Const POS_ERR_NUMBER As Long = 9  ' Subscript out of range.
	Const KEY_ERR_NUMBER As Long = 5  ' Invalid procedure call or argument.
	
	On Error GoTo ITEM_ERROR
	clx.Item index
	
	Clx_Has = True
	Exit Function
	
ITEM_ERROR:
	If VBA.Err.Number = POS_ERR_NUMBER Or VBA.Err.Number = KEY_ERR_NUMBER Then
		Clx_Has = False
	Else
		Err_Raise VBA.Err
	End If
End Function


' Get an item (safely) from a Collection.
Private Function Clx_Get(ByRef clx As Collection, _
	ByVal index As Variant, _
	Optional ByRef has As Boolean _
) As Variant
	has = Clx_Has(clx, index)
	
	If has Then
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
	
	clx.Add val, key := key
End Sub


' Get the length (along a dimension) of an array.
Private Function Arr_Length(ByRef arr As Variant, _
	Optional ByVal dmn As Long = 1 _
) As Long
	Const EMPTY_ERR_NUMBER As Long = 9  ' Subscript out of range.
	
	On Error GoTo BOUND_ERROR
	Arr_Length = UBound(arr, dmn) - LBound(arr, dmn) + 1
	Exit Function
	
BOUND_ERROR:
	If VBA.Err.Number = EMPTY_ERR_NUMBER Then
		Arr_Length = 0
	Else
		Err_Raise VBA.Err
	End If
End Function


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


' Test if text contains a substring.
Private Function Txt_Contains(ByVal txt As String, _
	ByVal sbs As String _
) As Boolean
	Const IDX_NONE As Long = 0
	
	Txt_Contains = (VBA.InStr(txt, sbs) <> IDX_NONE)
End Function
