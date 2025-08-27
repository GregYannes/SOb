# Utilities #

## Description ##

These procedures support **`SOb`** and are handy for your general use.

  - `Assign()` assigns any value (scalar or [objective][vba_isobj]) to your variable by [reference][vba_byref].
  - `Txt_Indent()` indents some text.
  - `Txt_Contains()` detects a substring within some text.

> [!NOTE]
> 
> Text detection is [case-sensitive][vba_txt_cmp].

  - `Clx_Has()` detects an [`.Item`][vba_clx_itm] within a [`Collection`][vba_clx].
  - `Clx_Get()` safely retrieves an [`.Item`][vba_clx_itm] from a `Collection`.
  - `Clx_Set()` sets the value of an `.Item` in a `Collection`.
  - `Arr_Length()` gets the length of an [array][vba_arr].
  - `Err_Raise()` conveniently [propagates][vba_ppg_err] the latest [`Err` object][vba_err_obj].


## Syntax ##

These procedures have the following syntax.

```vba
Assign var, val

Txt_Indent(txt, [indent], [before])

Txt_Contains(txt, substring)

Clx_Has(clx, index)

Clx_Get(clx, index, [has])

Clx_Set clx, key, val

Arr_Length(arr, [dimension])

Err_Raise [e]
```

They have the following named parameters.

| Name        | Type                      | Required | Default                     | Description                                                                                                |
| :---------- | :------------------------ | :------: | :-------------------------- | :--------------------------------------------------------------------------------------------------------- |
| `var`       | [^1]                      | ✓        |                             | The [variable][vba_vrb] to which `val` should be assigned[^2].                                             |
| `val`       | [^1]                      | ✓        |                             | The value you wish to assign.                                                                              |
| `txt`       | `String`                  | ✓        |                             | Some text.                                                                                                 |
| `indent`    | `String`                  |          | [`vbTab`][vba_tab]          | The spacing used to indent `txt`.  Defaults to a standard [horizontal tab][hrz_tab] like most indentation. |
| `before`    | `Boolean`                 |          | `True`                      | Indent (`True`) the first line of `txt`?                                                                   |
| `substring` | `String`                  | ✓        |                             | The substring to seek within `txt.                                                                         |
| `clx`       | [`Collection`][vba_clx]   | ✓        |                             | Any `Collection`.                                                                                          |
| `index`     | `Long`<br><br>`String`    | ✓        |                             | The position (`Long`) or key (`String`) of the [`.Item`][vba_clx_itm] in `clx`.                            |
| `has`       | `Boolean`                 |          |                             | A flag[^2] variable to track whether `clx` actually _has_ (`True`) an `.Item` at `index`.                  |
| `key`       | `String`                  | ✓        |                             | The `key` of the `.Item` in `clx`, whose `val`ue you wish to set.                                          |
| `arr`       | [Array][vba_arr]          | ✓        |                             | Any array.                                                                                                 |
| `dimension` | `Long`                    |          | `1`                         | The [dimension][vba_arr_dmn] along which you wish to measure `arr`.                                        |
| `e`         | `ErrObject`[^3]           |          | [`Err`][vba_err_obj] object | The latest error thrown during execution.                                                                  |


> [!NOTE]
> 
> Be sure to use the [`vbNewLine`][vba_newln] for line breaks, when you assemble (say) `txt` and other such text.  This uses the newline [specific to the system][sys_newln], and ensures that [`Obj_FormatFields*()`][sob_vis] and `Txt_Indent()` work as expected.


  [^1]: While this is technically a [`Variant`][vba_var], it accommodates any `var`iable or `val`ue you desire.
  [^2]: The procedure updates this variable by [reference][vba_byref], which overwrites any value it originally had.
  [^3]: The `ErrObject` is not a traditional "type", since there is only [one (global) instance][vba_err_typ] of the `Err` object.


## Output ##

These procedures have the following output.

  - `Assign()` returns no value, but rather assigns the `val`ue to your `var`iable by [reference][vba_byref].
  - `Txt_Indent()` returns a `String` where each line of `txt` is indented with `indent`.
  - `Txt_Contains()` returns `True` when `txt` contains the `substring`, and `False` otherwise.
  - `Clx_Has()` returns `True` when `clx` has an [`.Item`][vba_clx_itm] at `index`, and `False` otherwise.
  - `Clx_Get()` returns a `Variant` with the value at `index`, or swallows the error when no such `.Item` exists.
  - `Clx_Set()` returns no value.  It assigns the `val`ue to any `.Item` at `key`, or [adds][vba_clx_add] it when it does not already exist.
  - `Arr_Length()` returns a `Long` with the length of `arr`, as measured along the given `dimension`.
  - `Err_Raise()` returns no value.  It [raises][vba_err_rse] the latest [`Err`][vba_err_obj]or without forcing you to specify its details.


## Details ##

![](../med/banner_unfinished.png)


## Examples ##

### Assignment ###

Assign scalar values to [variables][vba_vrb]…

```vba
Dim sVar As String, vVar As Variant

Assign sVar, "first text"
Debug.Print sVar

Assign vVar, sVar
Debug.Print vVar

Assign vVar, "second text"
Debug.Print vVar

Assign sVar, vVar
Debug.Print sVar
```

> ```
> first text
> first text
> second text
> second text
> ```

<br>

…and [objective][vba_isobj] values like [`Range`][vba_rng]s[^4].

```vba
Dim rVar As Range, oVar As Object

Assign rVar, [A1:B2]
Debug.Print rVar.Address

Assign vVar, rVar
Debug.Print vVar.Address

Assign oVar, vVar
Debug.Print oVar.Address
```

> ```
> $A$1:$B$2
> $A$1:$B$2
> $A$1:$B$2
> ```


### Text Indentation ###

Indent some text.

```vba
Dim text As String
text = "First line."  & vbNewLine & _
       "Second line." & vbNewLine & _
       "Third line."

Debug.Print Txt_Indent(text)
Debug.Print

Debug.Print Txt_Indent(text, before := False)
Debug.Print

Debug.Print Txt_Indent(text, indent := "--> ")
```

> ```
> 	First line.
> 	Second line.
> 	Third line.
> 
> First line.
> 	Second line.
> 	Third line.
> 
> --> First line.
> --> Second line.
> --> Third line.
> ```


### Text Detection ###

Detect some text.

```vba
Debug.Print Txt_Contains(text, "First")
Debug.Print Txt_Contains(text, "FIRST")
Debug.Print Txt_Contains(text, "Fourth")
```

> ```
> True
> False
> False
> ```


### Manipulate `Collection`s ###

Detect an [`.Item`][vba_clx_itm] within a [`Collection`][vba_clx]…

```vba
Dim clx As Collection: Set clx = New Collection
clx.Add 10, key := "first"

Debug.Print Clx_Has(clx, 1)
Debug.Print Clx_Has(clx, "first")

Debug.Print Clx_Has(clx, 2)
Debug.Print Clx_Has(clx, "second")
```

> ```
> True
> True
> False
> False
> ```

<br>

…and get its value…

```vba
Dim flag As Boolean: flag = False

Debug.Print Clx_Get(clx, 1)
Debug.Print Clx_Get(clx, "first", has := flag)
Debug.Print flag

Debug.Print Clx_Get(clx, 2)
Debug.Print Clx_Get(clx, "second", has := flag)
Debug.Print flag
```

> ```
> 10
> 10
> True
> 
> 
> False
> ```

<br>

…and set its value.

```vba
Clx_Set clx, "first", -1
Debug.Print clx.Item("first")

Clx_Set clx, "second", 20
Debug.Print clx.Item("second")
```

> ```
> -1
> 20
> ```


### Array Length ###

Measure an [array][vba_arr].

```vba
Debug.Print "Declaring..."
Dim arr() As Variant
Debug.Print Arr_Length(arr)

Debug.Print "Initializing..."
ReDim arr(1 To 2, 0 To 3)
Debug.Print Arr_Length(arr)
Debug.Print Arr_Length(arr, dimension := 2)
```

> ```
> Declaring...
> 0
> Initializing...
> 2
> 4
> ```


### Error Propagation ###

[Propagate][vba_ppg_err] the latest [error][vba_err_obj].

```vba
	Debug.Print "Catching..."
	On Error GoTo PROPAGATE
	
	Dim num As Integer: num = "Text"
	Debug.Print "Succeeding..."
	
PROPAGATE:
	Debug.Print "Propagating..."
	Err_Raise
```

> ```
> Catching...
> Propagating...
> ```
> ![][vbe_err_ex]


  [^4]: You may specify a [`Range`][vba_rng] with its [`.Address`][vba_rng_adr] in [shortcut notation][vba_sct_nt]: `[A1:B2]`.


## See Also ##

Topics in this project…

  - [`Obj_FormatFields()`][sob_vis]
  - [`Obj_FormatFields0()`][sob_vis]

…and in VBA…

  - [`IsObject()`][vba_isobj]
  - Passing [`ByRef`][vba_byref]erence
  - [Case sensitivity][vba_txt_cmp]
  - [`Collection`][vba_clx]s
  - [`.Item()`][vba_clx_itm] method
  - [Arrays][vba_arr]
  - [Error propagation][vba_ppg_err]
  - [`Err`][vba_err_obj] object
  - [Variables][vba_vrb]
  - [`vbTab`][vba_tab]
  - Array [dimensions][vba_arr_dmn]
  - [`vbNewLine`][vba_newln]
  - [`Variant`][vba_var]s
  - [`.Add()`][vba_clx_add] method
  - [`.Raise()`][vba_err_rse] method
  - [`Range`][vba_rng]s
  - [`.Address`][vba_rng_adr] property
  - [Shortcut notation][vba_sct_nt]
  - [Error messages][vba_errs]

…and elsewhere.

  - [Horizontal tab][hrz_tab]
  - [System newline][sys_newln]



  [vba_isobj]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/isobject-function
  [vba_byref]:   https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [vba_txt_cmp]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/instr-function#settings
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_clx_itm]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_ppg_err]: https://www.fastercapital.com/content/Error-Handling--Error-Handling-Excellence--Bulletproofing-Your-VBA-Code.html#Error-Bubbling-and-Propagation
  [vba_err_obj]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
  [vba_vrb]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/declaring-variables
  [vba_tab]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/miscellaneous-constants
  [hrz_tab]:     https://www.ascii-code.com/9
  [vba_arr_dmn]: https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/arrays/array-dimensions
  [vba_newln]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/miscellaneous-constants
  [sys_newln]:   https://learn.microsoft.com/dotnet/api/system.environment.newline?view=net-9.0#property-value
  [sob_vis]:     Visualization.md
  [vba_var]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [vba_err_typ]: https://stackoverflow.com/a/55067026
  [vba_clx_add]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/add-method-visual-basic-for-applications
  [vba_err_rse]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/raise-method
  [vba_rng]:     https://learn.microsoft.com/office/vba/api/excel.range(object)
  [vbe_err_ex]:  ../med/vbe_error_13.png
  [vba_rng_adr]: https://learn.microsoft.com/office/vba/api/excel.range.address
  [vba_sct_nt]:  https://learn.microsoft.com/office/vba/excel/concepts/cells-and-ranges/refer-to-cells-by-using-shortcut-notation
  [vba_errs]:    https://learn.microsoft.com/office/vba/language/reference/error-messages
