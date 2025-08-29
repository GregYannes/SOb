# Visualization #

## Description ##

These functions support ["pretty printing"][pprint] for SObs.

  - `Obj_Format()` formats a textual representation of an SOb.
  - `Obj_Print()` also [prints][vba_print] that format to the [console][vbe_immed].
  - [`Obj_Print0()`][sob_fn0] is a primitive form of `Obj_Print()`, which prints to the console without formatting.
  - `Obj_FormatFields()` formats a textual representation of the [fields][sob_fld] in an SOb, which you supply in a [single array][vba_arr].  This representation is often used for the `details` of an SOb.
  - [`Obj_FormatFields0()`][sob_fn0] is a convenient form of `Obj_FormatFields()`, where you supply the fields directly.


## Syntax ##

These functions have the following syntax.

```vba
Obj_Format(obj, [depth], [plain], [pointer], [summary], [details], [preview], [orphan], [indent], [break])

Obj_Print(obj, [depth], [plain], [pointer], [summary], [details], [preview], [orphan], [indent], [break])

Obj_Print0([format])

Obj_FormatFields(fields, [separator])

Obj_FormatFields0(…)
```

They have the following named parameters.

| Name        | Type                              | Required | Default                  | Description                                                                                                                                                                                 |
| :---------- | :-------------------------------- | :------: | :----------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `obj`       | [`Collection`][vba_clx]           | ✓        |                          | An SOb you wish to visualize.                                                                                                                                                               |
| `depth`     | `Integer`                         |          | `1`                      | The depth to which the visualization should expand: `0` shows a `summary`, while positive expands to any `details`.                                                                         |
| `plain`     | `Boolean`                         |          | `False`                  | Format the SOb in plain format (`True`) rather than rich (`False`)?  See [**Examples**][sob_vis_ex] for appearance.                                                                         |
| `pointer`   | `Boolean`                         |          | `False`                  | Fall back to showing the [pointer][vba_ptr] (`True`) for `obj`, when we have neither `summary` nor `preview`?                                                                               |
| `summary`   | `String`                          |          | `""`                     | An expression summarizing the contents of `obj` on a single line, when `depth := 0`.  See `preview` and `pointer` for fallbacks when `summary := ""`.                                       |
| `details`   | `String`                          |          | `""`                     | Expressions detailing the contents of `obj` across multiple lines, when `depth > 0`.  See [**Examples**][sob_vis_ex] for formatting with `Obj_FormatFields*()`.                             |
| `preview`   | `Boolean`                         |          | `False`                  | Fall back to showing a preview of the `details`, when we have no `summary`?  See [**Examples**][sob_vis_ex] for appearance.                                                                 |
| `orphan`    | `Boolean`                         |          | `True`                   | Should a single line of `details` still be nested (`True`) or remain on a single line (`False`)?                                                                                            |
| `indent`    | `String`                          |          | [`vbTab`][vba_tab]       | The indentation used for nesting `details`.  Defaults to a standard [horizontal tab][hrz_tab] like most indentation.                                                                        |
| `break`     | `String`                          |          | [`vbNewLine`][vba_newln] | The linebreak which identifies where lines in `details` should be indented.  See [`Txt_Indent()`][sob_utils] for details on [usage][sob_brk_use].                                           |
| `format`    | `String`                          |          | `""`                     | Output for the console, which should already be formatted as desired.                                                                                                                       |
| `fields`    | [Array][vba_arr] of `String`s[^1] | ✓        |                          | An array with pairs of (textual) expressions: a field name followed by its value.  See [**Examples**][sob_vis_ex] for appearance.<br><br>This is best achieved via [`Array()`][vba_arr_fn]. |
| `separator` | `String`                          |          | `vbNewLine`              | The textual separator displayed between each pairing and the next.  Defaults to the [system newline][sys_newln], so each pair (`.field = value`) gets its own line.                         |
| …[^2]       | `String`s                         |          |                          | The pairs themselves, entered as individual arguments.<br><br>This is technically a [`ParamArray`][vba_parr].                                                                               |


> [!TIP]
> 
> Use [`vbNewLine`][vba_newln] for line breaks, when you assemble (say) `details` and other such text.  This uses the newline [specific to the system][sys_newln], and helps `Obj_FormatFields*()` and [`Txt_Indent()`][sob_utils] work frictionlessly.


## Output ##

These functions have the following output.

  - `Obj_Format()` returns a `String` with the formatted representation of `obj`.
  - `Obj_Print()` returns the same `String` and also [prints][vba_print] it to the [console][vbe_immed].
  - `Obj_Print0()` returns `format` and also prints it to the console "as is".
  - `Obj_FormatFields()` returns a `String` with the formatted representation (`.field = value`) of all pairs in `fields`.
  - `Obj_FormatFields0()` does likewise for all pairs in `…`.


## Examples ##

### Field Formatting ###

Format some [field][sob_fld] expressions programmatically…

```vba
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
```

> ```
> .Bar = 10
> .Baz = "Twenty"
> .Qux = [$A$1:$B$2]
> 
> .Bar = 10, .Baz = "Twenty", .Qux = [$A$1:$B$2]
> ```

<br>

…and manually.

```vba
fFormat = Obj_FormatFields0( _
	"Bar", "10", _
	"Baz", """Twenty""", _
	"Qux", "[$A$1:$B$2]" _
)

Debug.Print fFormat
```

> ```
> .Bar = 10
> .Baz = "Twenty"
> .Qux = [$A$1:$B$2]
> ```


### Printing Objects ###

Print a **"Foo"** object in `detail`…

```vba
Dim foo1 As Object: Set foo1 = New_Obj("Foo")

Obj_Print foo1, depth := 1, details := fFormat
```

> ```
> <Foo: {
> 	.Bar = 10
> 	.Baz = "Twenty"
> 	.Qux = [$A$1:$B$2]
> }>
> ```

<br>

…and in `summary`.

```vba
Obj_Print foo1, depth := 0, summary := "3"
```

> ```
> <Foo[3]>
> ```


### Summarizing Objects ###

In the absence of a `summary`, default to a `preview` of any `details`…

```vba
Obj_Print foo1, depth := 0, preview := True
Debug.Print

Obj_Print foo1, depth := 0, preview := True, details := fFormat
```

> ```
> <Foo: {}>
> 
> <Foo: {…}>
> ```

<br>

…or to the `pointer`…

```vba
Obj_Print foo1, depth := 0, pointer := True
```

> ```
> <Foo @105553142930832>
> ```

<br>

…or keep it simple.

```vba
Obj_Print foo1, depth := 0
```

> ```
> <Foo>
> ```


### Plain Formatting ###

Format plainly in `detail`…

```vba
Obj_Print foo1, depth := 1, plain := True
Debug.Print

Obj_Print foo1, depth := 1, plain := True, details := fFormat
```

> ```
> {}
> 
> {
> 	.Bar = 10
> 	.Baz = "Twenty"
> 	.Qux = [$A$1:$B$2]
> }
> ```

<br>

…and in `summary`.

```vba
Obj_Print foo1, depth := 0, plain := True
Debug.Print

Obj_Print foo1, depth := 0, plain := True, details := fFormat
```

> ```
> {}
> 
> {…}
> ```


### Miscellaneous Settings ###

Play with orphan lines…

```vba
Obj_Print foo1, depth := 1, details := ".Bar = 10", orphan := False
Debug.Print

Obj_Print foo1, depth := 1, details := ".Bar = 10", orphan := True
```

> ```
> <Foo: {.Bar = 10}>
> 
> <Foo: {
> 	.Bar = 10
> }>
> ```

<br>

…and with indentation…

```vba
Obj_Print foo1, depth := 1, details := fFormat, indent := "--> "
```

> ```
> <Foo: {
> --> .Bar = 10
> --> .Baz = "Twenty"
> --> .Qux = [$A$1:$B$2]
> }>
> ```

<br>

…at various breaks.

```vba
Obj_Print foo1, depth := 1, details := fFormat, indent := " -->", break := "="
```

> ```
> <Foo: {
> .Bar = --> 10
> .Baz = --> "Twenty"
> .Qux = --> [$A$1:$B$2]
> }>
> ```


### Implement Printing ###

[Define][vba_enum] a few [fields][sob_fld] for an SOb of class **"Foo"**…

```vba
Enum Foo__Fields
	Bar
	Baz
	Qux
End Enum
```

<br>

…along with their [accessors][sob_tmp_acc].

```vba
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
```

<br>

Define[^3] simple [`Foo_Format()`][sob_tmp_fmt] and [`Foo_Print()`][sob_tmp_prn] methods for **"Foo"**.

```vba
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
		indent := vbTab, _
		break := vbNewLine _
	)
End Function


Function Foo_Print(foo As Object, _
	Optional depth As Integer = 1, _
	Optional plain As Boolean = False _
) As String
	Foo_Print = Foo_Format(foo, depth := depth, plain := plain)
	
	Obj_Print0 Foo_Print
End Function
```

<br>

Print a **"Foo"** object[^4] richly…

```vba
Foo_Bar(foo1) = -1
Foo_Baz(foo1) = "text"
Set Foo_Qux(foo1) = [C3:D4]

Foo_Print foo1, depth := 0
Debug.Print

Foo_Print foo1, depth := 1
```

> ```
> <Foo[3]>
> 
> <Foo: {
> 	.Bar = -1
> 	.Baz = "text"
> 	.Qux = [$C$3:$D$4]
> }>
> ```

<br>

…and plainly.

```vba
Foo_Print foo1, depth := 0, plain := True
Debug.Print

Foo_Print foo1, depth := 1, plain := True
```

> ```
> {…}
> 
> {
> 	.Bar = -1
> 	.Baz = "text"
> 	.Qux = [$C$3:$D$4]
> }
> ```


### Nested Printing ###

[Define][vba_enum] a few [fields][sob_fld] for an SOb of class **"Snaf"**…

```vba
Enum Snaf__Fields
	TxtField
	FooField
End Enum
```

<br>

…along with their [accessors][sob_tmp_acc].

```vba
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
```

<br>

Define a simple [`Snaf_Print()`][sob_tmp_prn] method for **"Snaf"**, which decrements `depth` and passes it (and `plain`) recursively to the `Foo_Print()` method for its `FooField` field.

```vba
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
		indent := vbTab, _
		break := vbNewLine _
	)
	
	Obj_Print0 Snaf_Print
End Function
```

<br>

Print a **"Snaf"** object to various `depth`s, in rich format…

```vba
Dim snaf1 As Object: Set snaf1 = New_Obj("Snaf")

Snaf_TxtField(snaf1) = "some more text"
Set Snaf_FooField(snaf1) = foo1

Snaf_Print snaf1, depth := 0
Debug.Print

Snaf_Print snaf1, depth := 1
Debug.Print

Snaf_Print snaf1, depth := 2
```

> ```
> <Snaf[Situation normal]>
> 
> <Snaf: {
> 	.TxtField = "some more text"
> 	.FooField = <Foo[3]>
> }>
> 
> <Snaf: {
> 	.TxtField = "some more text"
> 	.FooField = <Foo: {
> 		.Bar = -1
> 		.Baz = "text"
> 		.Qux = [$C$3:$D$4]
> 	}>
> }>
> ```

<br>

…and in plain format.

```vba
Snaf_Print snaf1, depth := 0, plain := True
Debug.Print

Snaf_Print snaf1, depth := 1, plain := True
Debug.Print

Snaf_Print snaf1, depth := 2, plain := True
```

> ```
> {…}
> 
> {
> 	.TxtField = "some more text"
> 	.FooField = {…}
> }
> 
> {
> 	.TxtField = "some more text"
> 	.FooField = {
> 		.Bar = -1
> 		.Baz = "text"
> 		.Qux = [$C$3:$D$4]
> 	}
> }
> ```


## See Also ##

Topics in this project…

  - [`*0()`][sob_fn0] family
  - [Fields][sob_fld]
  - [`Txt_Indent()`][sob_utils]
  - [Templates][sob_tmps]
  - [Setup][sob_setup] with templates
  - [Enumerated fields][sob_tmp_enm]
  - [Field accessors][sob_tmp_acc]
  - [Formatting][sob_tmp_fmt]
  - [Printing][sob_tmp_prn]

…in VBA…

  - [`.Print()`][vba_print] method
  - [Immediate window][vbe_immed] in the [Visual Basic Editor][vbe] (VBE)
  - [Arrays][vba_arr]
  - [`Collection`][vba_clx]s
  - [Pointers][vba_ptr]
  - [`vbTab`][vba_tab]
  - [`Array()`][vba_arr_fn]
  - [`vbNewLine`][vba_newln]
  - [`ParamArray`][vba_parr]s
  - [Named arguments][vba_nm_args]
  - [`Enum`][vba_enum]erations
  - [`CStr()`][vba_cstr]
  - [`Range`][vba_rng]s
  - [`.Address`][vba_rng_adr] property
  - [Shortcut notation][vba_sct_nt]

…and elsewhere.

  - [Pretty-printing][pprint]
  - [Horizontal tab][hrz_tab]
  - [System newline][sys_newln]



  [^1]: You may use a `String()` array, or a `Variant()` array containing `String`s.
  [^2]: [`ParamArray`][vba_parr]s like `…` are not actually passed to a single [named argument][vba_nm_args], but rather as several nameless arguments.
  [^3]: Use [`CStr()`][vba_cstr] to convert various values into textual `String`s.
  [^4]: You may specify a [`Range`][vba_rng] with its [`.Address`][vba_rng_adr] in [shortcut notation][vba_sct_nt]: `[A1:B2]`.



  [pprint]:      https://en.wikipedia.org/wiki/Pretty-printing
  [vba_print]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/print-method
  [vbe_immed]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/immediate-window
  [sob_fn0]:     Zero.md
  [sob_fld]:     Field.md
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [sob_vis_ex]:  #examples
  [vba_ptr]:     https://classicvb.net/tips/varptr
  [vba_tab]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/miscellaneous-constants
  [hrz_tab]:     https://www.ascii-code.com/9
  [vba_newln]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/miscellaneous-constants
  [sob_utils]:   Utilities.md
  [sob_brk_use]: Utilities.md#text-indentation
  [vba_arr_fn]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/array-function
  [sys_newln]:   https://learn.microsoft.com/dotnet/api/system.environment.newline?view=net-9.0#property-value
  [vba_parr]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-parameter-arrays
  [vba_enum]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_tmp_acc]: ../src/SObTemplate.bas#L171-L213
  [sob_tmp_fmt]: ../src/SObTemplate.bas#L277-L302
  [sob_tmp_prn]: ../src/SObTemplate.bas#L254-L273
  [sob_tmps]:    ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:   Setup.md
  [sob_tmp_enm]: ../src/SObTemplate.bas#L26-L29
  [vbe]:         https://learn.microsoft.com/office/vba/library-reference/concepts/getting-started-with-vba-in-office#macros-and-the-visual-basic-editor
  [vba_nm_args]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
  [vba_cstr]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/type-conversion-functions
  [vba_rng]:     https://learn.microsoft.com/office/vba/api/excel.range(object)
  [vba_rng_adr]: https://learn.microsoft.com/office/vba/api/excel.range.address
  [vba_sct_nt]:  https://learn.microsoft.com/office/vba/excel/concepts/cells-and-ranges/refer-to-cells-by-using-shortcut-notation
