Attribute VB_Name = "Ex_Utilities"



' #############
' ## Options ##
' #############

Option Compare Binary



' ##############
' ## Examples ##
' ##############

Public Sub Utilities()
	Debug.Print "###############"
	Debug.Print "## Utilities ##"
	Debug.Print "###############"
	
	Debug.Print
	Debug.Print "############################"
	Debug.Print "## Utilities | Assignment ##"
	Debug.Print "############################"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_ASSIGNMENT_1_START:
Dim sVar As String, vVar As Variant

Assign sVar, "first text"
Debug.Print sVar

Assign vVar, sVar
Debug.Print vVar

Assign vVar, "second text"
Debug.Print vVar

Assign sVar, vVar
Debug.Print sVar
	UTILITIES_ASSIGNMENT_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_ASSIGNMENT_2_START:
Dim rVar As Range, oVar As Object

Assign rVar, [A1:B2]
Debug.Print rVar.Address

Assign vVar, rVar
Debug.Print vVar.Address

Assign oVar, vVar
Debug.Print oVar.Address
	UTILITIES_ASSIGNMENT_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "##################################"
	Debug.Print "## Utilities | Text Indentation ##"
	Debug.Print "##################################"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_TEXT_INDENTATION_1_START:
Dim text As String
text = "First line."  & vbNewLine & _
       "Second line." & vbNewLine & _
       "Third line."

Debug.Print Txt_Indent(text)
Debug.Print

Debug.Print Txt_Indent(text, before := False)
Debug.Print

Debug.Print Txt_Indent(text, indent := "--> ")
Debug.Print

Debug.Print Txt_Indent(text, indent := "--> ", break := " ", before := False)
	UTILITIES_TEXT_INDENTATION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "###############################"
	Debug.Print "## Typology | Text Detection ##"
	Debug.Print "###############################"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_TEXT_DETECTION_1_START:
Debug.Print Txt_Contains(text, "First")
Debug.Print Txt_Contains(text, "Second", start := 13)
Debug.Print Txt_Contains(text, "Second", start := 14)
Debug.Print Txt_Contains(text, "THIRD")
Debug.Print Txt_Contains(text, "THIRD", compare := vbTextCompare)
Debug.Print Txt_Contains(text, "Fourth")
	UTILITIES_TEXT_DETECTION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "#######################################"
	Debug.Print "## Typology | Manipulate Collections ##"
	Debug.Print "#######################################"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_MANIPULATE_COLLECTIONS_1_START:
Dim clx As Collection: Set clx = New Collection
clx.Add 10, key := "first"

Debug.Print Clx_Has(clx, 1)
Debug.Print Clx_Has(clx, "first")

Debug.Print Clx_Has(clx, 2)
Debug.Print Clx_Has(clx, "second")
	UTILITIES_MANIPULATE_COLLECTIONS_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_MANIPULATE_COLLECTIONS_2_START:
Dim flag As Boolean: flag = False

Debug.Print Clx_Get(clx, 1)
Debug.Print Clx_Get(clx, "first", has := flag)
Debug.Print flag

Debug.Print Clx_Get(clx, 2)
Debug.Print Clx_Get(clx, "second", has := flag)
Debug.Print flag
	UTILITIES_MANIPULATE_COLLECTIONS_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_MANIPULATE_COLLECTIONS_3_START:
Clx_Set clx, "first", -1
Debug.Print clx.Item("first")

Clx_Set clx, "second", 20
Debug.Print clx.Item("second")
	UTILITIES_MANIPULATE_COLLECTIONS_3_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "#############################"
	Debug.Print "## Typology | Array Length ##"
	Debug.Print "#############################"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_ARRAY_LENGTH_1_START:
Debug.Print "Declaring..."
Dim arr() As Variant
Debug.Print Arr_Length(arr)

Debug.Print "Initializing..."
ReDim arr(1 To 2, 0 To 3)
Debug.Print Arr_Length(arr)
Debug.Print Arr_Length(arr, dimension := 2)
	UTILITIES_ARRAY_LENGTH_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "##################################"
	Debug.Print "## Typology | Error Propagation ##"
	Debug.Print "##################################"
	
	
	Debug.Print
	Debug.Print "```"
	UTILITIES_ERROR_PROPAGATION_1_START:
	Debug.Print "Catching..."
	On Error GoTo PROPAGATE
	
	Dim num As Integer: num = "Text"
	Debug.Print "Succeeding..."
	
PROPAGATE:
	Debug.Print "Propagating..."
	Err_Raise
	UTILITIES_ERROR_PROPAGATION_1_END:
	Debug.Print "```"
End Sub
