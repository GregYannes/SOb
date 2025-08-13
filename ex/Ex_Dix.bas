Attribute VB_Name = "Ex_Dix"



' ###############
' ## Constants ##
' ###############

' Class name for the simulated "Dix" object.
Private Const DIX_CLS As String = "Dix"





' ...





' ##################
' ## Enumerations ##
' ##################

' Fields for the simulated "Dix" object.
Private Enum Dix_Field
	Keys
	Items
	Count
End Enum





' ...





' #########
' ## Dix ##
' #########

' ####################
' ## Dix | Creation ##
' ####################

' Construct a simulated "Dix" object.
Public Function New_Dix() As Object
	Dix_Initialize New_Dix
End Function



' Initialize a simulated "Dix" object.
Public Sub Dix_Initialize(ByRef dix As Object)
	Const CLS_NAME As String = DIX_CLS
	
	SOb.Obj_Initialize dix, CLS_NAME
	
	
	If Not SOb.Obj_HasField(dix, Dix_Field.Keys) Then
		Dim keys As Collection: Set keys = New Collection
		Set Dix_Keys(dix) = keys
	End If
	
	If Not SOb.Obj_HasField(dix, Dix_Field.Items) Then
		Dim items As Collection: Set items = New Collection
		Set Dix_Items(dix) = items
	End If
	
	If Not SOb.Obj_HasField(dix, Dix_Field.Count) Then
		Dim count As Long: count = Dix_Keys(dix).Count
		Let Dix_Count(dix) = count
	End If
	
	' ...
End Sub



' ####################
' ## Dix | Typology ##
' ####################

' Identify a simulated "Dix" object.
Public Function IsDix(ByRef x As Variant, _
	Optional ByVal strict As Boolean = False _
) As Boolean
	Const CLS_NAME As String = DIX_CLS
	
	
	' ### Class and Fields ###
	
	' Ensure an accurate class with its proper set of fields.
	IsDix = SOb.IsObj(x, class := CLS_NAME, strict := strict, fields := Array( _
		Dix_Field.Keys, _
		Dix_Field.Items, _
		Dix_Field.Count _
		 _
		 _
		 _
	))
	If Not IsDix Then Exit Function
	
	
	' ### Accessors ###
	
	' Treat as an object moving forward.
	Dim obj As Object: Set obj = x
	
	' Ensure the field accessors all work.
	On Error GoTo CHECK_ERROR
	
	If IsDix Then SOb.Obj_Check _
		Dix_Keys(obj), _
		Dix_Items(obj), _
		Dix_Count(obj) _
		 _
		 _
		 _
	
	On Error GoTo 0
	If Not IsDix Then Exit Function
	
	
	' Return the result in lieu of errors.
	Exit Function
	
' Handle inaccessibility.
CHECK_ERROR:
	IsDix = SOb.Obj_CheckError(type_ := True)
End Function



' Cast to a simulated "Dix" object.
Public Function AsDix(ByRef x As Variant) As Object
	' Cast the input to a (generic) simulated object...
	Dim obj As Object: Set obj = SOb.AsObj(x)
	
	' ...and extract its fields into a new "Dix" object.
	Set AsDix = New_Dix()
	
	Set Dix_Keys(AsDix) = Dix_Keys(obj)
	Set Dix_Items(AsDix) = Dix_Items(obj)
	Let Dix_Count(AsDix) = Dix_Count(obj)
End Function



' ##################
' ## Dix | Fields ##
' ##################

' The keys to the items in the dictionary.
Private Property Get Dix_Keys(ByRef dix As Object) As Collection
	SOb.Obj_Get Dix_Keys, dix, Dix_Field.Keys
End Property

Private Property Set Dix_Keys(ByRef dix As Object, ByVal val As Collection)
	Set SOb.Obj_Field(dix, Dix_Field.Keys) = val
End Property



' The items in the dictionary.
Private Property Get Dix_Items(ByRef dix As Object) As Collection
	SOb.Obj_Get Dix_Items, dix, Dix_Field.Items
End Property

Private Property Set Dix_Items(ByRef dix As Object, ByRef val As Collection)
	Set SOb.Obj_Field(dix, Dix_Field.Items) = val
End Property



' The count of items in the dictionary.
Public Property Get Dix_Count(ByRef dix As Object) As Long
	SOb.Obj_Get Dix_Count, dix, Dix_Field.Count
End Property

Private Property Let Dix_Count(ByRef dix As Object, ByVal val As Long)
	Let SOb.Obj_Field(dix, Dix_Field.Count) = val
End Property



' ###################
' ## Dix | Methods ##
' ###################

' ' An external method which your user may call for a return value.
' Public Function Dix_MethodOne(ByRef dix As Object, _
' 	 _
' 	 _
' 	 _
' ) As Integer
' 	' ...
' 	
' 	' Let MethodOne = ...
' End Function



' ' An internal method which your user may NOT call.
' Private Sub Dix_MethodTwo(ByRef dix As Object, _
' 	 _
' 	 _
' 	 _
' )
' 	' ...
' End Sub



' TODO: Procedures for any further methods you desire.
' ...



' #########################
' ## Dix | Visualization ##
' #########################

' Print a simulated "Dix" object.
Public Function Dix_Print(ByRef dix As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	Dix_Print = Dix_Format(dix, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		preview := preview, _
		indent := indent, _
		orphan := orphan _
	)
	
	SOb.Obj_Print0 Dix_Print
End Function



' Format a simulated "Dix" object for printing.
Public Function Dix_Format(ByRef dix As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	' Summarize the "Dix" object with its size...
	Dim sum As String: sum = Dix_Count(dix)
	
	' ...and detail it with a breakdown of its fields.
	Dim dtl As String: dtl = SOb.Obj_FormatFields0( _
		"Keys",  "Collection[" & Dix_Keys(dix).Count & "]", _
		"Items", "Collection[" & Dix_Items(dix).Count & "]", _
		"Count", Dix_Count(dix) _
	)
	
	' Pass the settings for formatting.
	Dix_Format = SOb.Obj_Format(dix, _
		sum := sum, _
		dtl := dtl, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		pvw := preview, _
		ind := indent, _
		orf := orphan _
	)
End Function
