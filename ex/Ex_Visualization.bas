Attribute VB_Name = "Ex_Visualization"



' ##################
' ## Enumerations ##
' ##################

Private Enum Foo__Fields
	Bar
	Baz
	Qux
End Enum

Private Enum Snaf__Fields
	TxtField
	FooField
End Enum



' ################
' ## Procedures ##
' ################

Property Get Foo_Bar(foo As Object) As Integer
	Obj_Get Foo_Bar, foo, Bar
End Property

Property Let Foo_Bar(foo As Object, val As Integer)
	Let Obj_Field(foo, Bar) = val
End Property


Property Get Foo_Baz(foo As Object) As String
	Obj_Get Foo_Baz, foo, Baz
End Property

Property Let Foo_Baz(foo As Object, val As String)
	Let Obj_Field(foo, Baz) = val
End Property


Property Get Foo_Qux(foo As Object) As Range
	Obj_Get Foo_Qux, foo, Qux
End Property

Property Set Foo_Qux(foo As Object, val As Range)
	Set Obj_Field(foo, Qux) = val
End Property


Function Foo_Format(foo As Object, _
	depth As Integer, _
	plain As Boolean _
) As String
	Dim sSummary As String: sSummary = VBA.CStr(Obj_FieldCount(foo))
	
	Dim sDetails As String: sDetails = Obj_FormatFields0( _
		"Bar", VBA.CStr(Foo_Bar(foo)), _
		"Baz", """" & Foo_Baz(foo) & """", _
		"Qux", "[" & Foo_Qux(foo).Address & "]" _
	)
	
	Foo_Format = Obj_Format(foo, _
		summary := sSummary, _
		details := sDetails, _
		depth := depth, _
		plain := plain, _
		orphan := False, _
		indent := vbTab _
	)
End Function


Function Foo_Print(foo As Object, _
	Optional depth As Integer = 1, _
	Optional plain As Boolean = False _
) As String
	Foo_Print = Foo_Format(foo, depth := depth, plain := plain)
	
	Obj_Print0 Foo_Print
End Function


Property Get Snaf_TxtField(snaf As Object) As String
	Obj_Get Snaf_TxtField, snaf, TxtField
End Property

Property Let Snaf_TxtField(snaf As Object, val As String)
	Let Obj_Field(snaf, TxtField) = val
End Property


Property Get Snaf_FooField(snaf As Object) As Object
	Obj_Get Snaf_FooField, snaf, FooField
End Property

Property Set Snaf_FooField(snaf As Object, val As Object)
	Set Obj_Field(snaf, FooField) = val
End Property


Function Snaf_Print(snaf As Object, _
	Optional depth As Integer = 1, _
	Optional plain As Boolean = False _
) As String
	Dim sSummary As String: sSummary = "Situation normal"
	
	Dim sDetails As String: sDetails = Obj_FormatFields0( _
		"TxtField", """" & Snaf_TxtField(snaf) & """", _
		"FooField", Foo_Format(Snaf_FooField(snaf), depth := depth - 1, plain := plain) _
	)
	
	Snaf_Print = Obj_Format(snaf, _
		summary := sSummary, _
		details := sDetails, _
		depth := depth, _
		plain := plain, _
		orphan := False, _
		indent := vbTab _
	)
	
	Obj_Print0 Snaf_Print
End Function



' ##############
' ## Examples ##
' ##############

Public Sub Visualization()
	Debug.Print "###################"
	Debug.Print "## Visualization ##"
	Debug.Print "###################"
	
	Debug.Print
	Debug.Print "######################################"
	Debug.Print "## Visualization | Field Formatting ##"
	Debug.Print "######################################"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_FIELD_FORMATTING_1_START:
Dim fFormat As String
Dim aFields As Variant: aFields = Array( _
	"Bar", "10", _
	"Baz", """Twenty""", _
	"Qux", "[$A$1:$B$2]" _
)

fFormat = Obj_FormatFields(aFields)
Debug.Print fFormat
Debug.Print

fFormat = Obj_FormatFields(aFields, separator := ", ")
Debug.Print fFormat
	VISUALIZATION_FIELD_FORMATTING_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_FIELD_FORMATTING_2_START:
fFormat = Obj_FormatFields0( _
	"Bar", "10", _
	"Baz", """Twenty""", _
	"Qux", "[$A$1:$B$2]" _
)

Debug.Print fFormat
	VISUALIZATION_FIELD_FORMATTING_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "######################################"
	Debug.Print "## Visualization | Printing Objects ##"
	Debug.Print "######################################"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_PRINTING_OBJECTS_1_START:
Dim foo1 As Object: Set foo1 = New_Obj("Foo")

Obj_Print foo1, depth := 1, details := fFormat
	VISUALIZATION_PRINTING_OBJECTS_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_PRINTING_OBJECTS_2_START:
Obj_Print foo1, depth := 0, summary := "3"
	VISUALIZATION_PRINTING_OBJECTS_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "#########################################"
	Debug.Print "## Visualization | Summarizing Objects ##"
	Debug.Print "#########################################"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_SUMMARIZING_OBJECTS_1_START:
Obj_Print foo1, depth := 0, preview := True
Debug.Print

Obj_Print foo1, depth := 0, preview := True, details := fFormat
	VISUALIZATION_SUMMARIZING_OBJECTS_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_SUMMARIZING_OBJECTS_2_START:
Obj_Print foo1, depth := 0, pointer := True
	VISUALIZATION_SUMMARIZING_OBJECTS_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_SUMMARIZING_OBJECTS_3_START:
Obj_Print foo1, depth := 0
	VISUALIZATION_SUMMARIZING_OBJECTS_3_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "######################################"
	Debug.Print "## Visualization | Plain Formatting ##"
	Debug.Print "######################################"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_PLAIN_FORMATTING_1_START:
Obj_Print foo1, depth := 1, plain := True
Debug.Print

Obj_Print foo1, depth := 1, plain := True, details := fFormat
	VISUALIZATION_PLAIN_FORMATTING_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_PLAIN_FORMATTING_2_START:
Obj_Print foo1, depth := 0, plain := True
Debug.Print

Obj_Print foo1, depth := 0, plain := True, details := fFormat
	VISUALIZATION_PLAIN_FORMATTING_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "############################################"
	Debug.Print "## Visualization | Miscellaneous Settings ##"
	Debug.Print "############################################"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_MISCELLANEOUS_SETTINGS_1_START:
Obj_Print foo1, depth := 1, details := ".Bar = 10", orphan := False
Debug.Print

Obj_Print foo1, depth := 1, details := ".Bar = 10", orphan := True
	VISUALIZATION_MISCELLANEOUS_SETTINGS_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_MISCELLANEOUS_SETTINGS_2_START:
Obj_Print foo1, depth := 1, details := fFormat, indent := "--> "
	VISUALIZATION_MISCELLANEOUS_SETTINGS_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "########################################"
	Debug.Print "## Visualization | Implement Printing ##"
	Debug.Print "########################################"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_IMPLEMENT_PRINTING_1_START:
Foo_Bar(foo1) = -1
Foo_Baz(foo1) = "text"
Set Foo_Qux(foo1) = [C3:D4]

Foo_Print foo1, depth := 0
Debug.Print

Foo_Print foo1, depth := 1
	VISUALIZATION_IMPLEMENT_PRINTING_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_IMPLEMENT_PRINTING_2_START:
Foo_Print foo1, depth := 0, plain := True
Debug.Print

Foo_Print foo1, depth := 1, plain := True
	VISUALIZATION_IMPLEMENT_PRINTING_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "#####################################"
	Debug.Print "## Visualization | Nested Printing ##"
	Debug.Print "#####################################"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_NESTED_PRINTING_1_START:
Dim snaf1 As Object: Set snaf1 = New_Obj("Snaf")

Snaf_TxtField(snaf1) = "some more text"
Set Snaf_FooField(snaf1) = foo1

Snaf_Print snaf1, depth := 0
Debug.Print

Snaf_Print snaf1, depth := 1
Debug.Print

Snaf_Print snaf1, depth := 2
	VISUALIZATION_NESTED_PRINTING_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VISUALIZATION_NESTED_PRINTING_2_START:
Snaf_Print snaf1, depth := 0, plain := True
Debug.Print

Snaf_Print snaf1, depth := 1, plain := True
Debug.Print

Snaf_Print snaf1, depth := 2, plain := True
	VISUALIZATION_NESTED_PRINTING_2_END:
	Debug.Print "```"
End Sub
