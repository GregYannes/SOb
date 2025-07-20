Attribute VB_Name = "Test"

' ###############
' ## Constants ##
' ###############

' Class name for simulated Dix object.
Const CLS_DIX As String = "Dix"



' ##################
' ## Enumerations ##
' ##################

' Field support for simulated Dix object.
Private Enum DixField
    keys
    items
    count
End Enum



' #############
' ## Testing ##
' #############

' .
Sub Test()
    ' Dim dix As Object:     Obj_Initialize dix, "Dix"
    ' Dim dix As Collection: Obj_Initialize dix, "Dix"
    ' Dim dix As Object:     Set dix = New_Obj("Dix")
    Dim dix As Object: Set dix = New_Dix()
    
    Obj_Field(dix, DixField.count) = 42
    
    
    Dim copy As String: Obj_ClassKey copy
    Debug.Print "Obj_ClassKey() = """ & copy & """"
    Debug.Print "Obj_HasClass(dix) = " & Obj_HasClass(dix)
    Debug.Print "Obj_Class(dix) = """ & Obj_Class(dix) & """"
    
    Debug.Print
    
    Debug.Print "IsObj(dix) = " & IsObj(dix)
    Debug.Print "IsObj(dix, """ & CLS_DIX & """) = " & IsObj(dix, CLS_DIX)
    Debug.Print "IsObj(dix, ""Other"") = " & IsObj(dix, "Other")
    
    Debug.Print
    
    Obj_FieldKey copy, DixField.count  ' obj := dix
    Debug.Print "Obj_FieldKey(DixField.Count) = """ & copy & """"
    Debug.Print "Obj_HasField(dix, DixField.Count) = " & Obj_HasField(dix, DixField.count)
    Debug.Print "Obj_Field(dix, DixField.Count) = " & Obj_Field(dix, DixField.count)
End Sub



' ...



' ###################
' ## Dixionary SOB ##
' ###################

' Constructor.
Public Function New_Dix() As Object
    Const CLS_NAME = CLS_DIX
    
    Dim dix As Object: Set dix = New_Obj(CLS_NAME)
    Dix_Initialize dix
    
    Set New_Dix = dix
End Function


' Initializer.
Private Sub Dix_Initialize(ByRef dix As Object)
    Const CLS_NAME = CLS_DIX
    Obj_Initialize dix, CLS_NAME
    
    If Not Obj_HasField(dix, DixField.keys) Then
        Dim keys As Collection: Set keys = New Collection
        Set Obj_Field(dix, DixField.keys) = keys
    End If
    
    If Not Obj_HasField(dix, DixField.items) Then
        Dim items As Collection: Set items = New Collection
        Set Obj_Field(dix, DixField.items) = items
    End If
End Sub


' ' Validator.
' Private Function Dix_Validate(ByRef dix As Object) Then
'     ' ...
' End Function



' ...



' ##########
' ## SOBs ##
' ##########

' ################
' ## SOBs ¥ API ##
' ################

' ###########################
' ## SOBs ¥ API ¥ Creation ##
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
' ## SOBs ¥ API ¥ Typing ##
' #########################

' Check for a simulated object.
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
	If IsObj And cls <> VBA.vbNullString Then
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
' ## SOBs ¥ API ¥ Fields ##
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
	' Obj_FieldCount = Math_Max(0, Obj_FieldCount)
End Property


' Check for a simulated field.
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
	Assign Obj_Field, obj.Item(key)  ' Clx_Get(obj, key)
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
' ## SOBs ¥ API ¥ Visualization ##
' ################################

' Format a simulated object for printing.
Private Function Obj_Format(ByRef obj As Collection, _
	Optional ByVal class As String = VBA.vbNullString, _
	Optional ByVal summary As String = VBA.vbNullString, _
	Optional ByVal detail As String = VBA.vbNullString, _
	Optional ByVal indent As String = VBA.vbNullString _
) As String
	' ...
End Function


' Format the summary descriptor(s) of an object.
Private Function Obj_Format_Summary(ByVal val As String) As String
	' ...
End Function


' Format the detailed contents of an object.
Private Function Obj_Format_Detail(ByVal val As String) As String
	' ...
End Function


' Format a (single) simulated field for (detailed) printing.
Private Function Obj_FormatField( _
	ByVal name As String, _
	ByVal val As String _
) As String
	' ...
End Function


' Format (a set of) simulated fields for (detailed) printing.
Private Function Obj_FormatFields(ParamArray fields() As Variant) As String
	' ...
End Function



' ####################
' ## SOBs ¥ Helpers ##
' ####################

' Check for a simulated class.
Private Function Obj_HasClass(ByRef obj As Collection) As Boolean
	Dim key As String: Obj_ClassKey key
	Obj_HasClass = Clx_Has(obj, key)
End Function


' Get the class of a simulated object.
Private Property Get Obj_Class(ByRef obj As Collection) As String
	Dim key As String: Obj_ClassKey key
	If Obj_HasClass(obj) Then
		Obj_Class = obj.Item(key)  ' Clx_Get(obj, key)
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
	
	
	Const DEF_CLS = Empty
	Const FLD_PFX = "Field_"
	Const KEY_SEP = "."
	
	
	Dim cls As String, secret As String, key As String
	' If Obj_HasClass(obj) Then
	' 	cls = Obj_Class(obj)
	' Else
		cls = DEF_CLS
	' End If
	
	Obj_Secret secret
	fld = fld + 1
	
	key = FLD_PFX & VBA.CStr(fld) & KEY_SEP & secret
	If cls <> Empty Then
		key = cls & KEY_SEP & key
	End If
	
	var = key
End Sub


' Securely obtain the key for a simulated class: "Class.xxx"
Private Sub Obj_ClassKey(ByRef var As String)
	Const CLS_PFX = "Class"
	Const KEY_SEP = "."
	
	Dim key As String, secret As String
	Obj_Secret secret
	key = CLS_PFX & KEY_SEP & secret
	
	var = key
End Sub


' Securely obtain the secret token for keys.
Private Sub Obj_Secret(ByRef var As String)
' 	Optional ByVal refresh As Boolean = False
	
	
	Const SEC_PFX = "x"
	Const REF_SEP = ""
	
	
	Static secret As String, isInit As Boolean
	
	If Not isInit Then  ' Or refresh
		Dim ref1 As New Collection, ref2 As New Collection
		secret = SEC_PFX & VBA.Hex(VBA.ObjPtr(ref1)) & REF_SEP & VBA.Hex(VBA.ObjPtr(ref2))
		
		isInit = True
	End If
	
	var = secret
End Sub



' ######################
' ## SOBs ¥ Utilities ##
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


' ' .
' Private Function Text_Indent(ByVal txt As String, _
' 	Optional ByVal indent As String = Constants.vbTab _
' ) As String
' ' 	Optional ByVal old As String = VBA.vbNullString _
' 
' 	' ...
' End Function
