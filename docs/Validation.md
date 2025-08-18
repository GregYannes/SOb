# Validation #

## Description ##

These procedures help you validate [advanced implementations][sob_tmp_chk] of [`Is*()`][sob_tmp_is], where invalid input would otherwise crash.  Such validation helps your `Is*()` function identify your SObs precisely, beyond the basics supported by [`IsObj()`][sob_typo].

  - `Check()` conveniently calls your [accessors][sob_tmp_acc], merely to check that they work.  This lets you ignore the [field][sob_fld] value[^1] itself.
  - `CheckError()` catches certain errors (like [type mismatch][vba_err_13]) that invalidate the check, but it [propagates][vba_ppg_err] all other errors.  Your `Is*` function should [return][sob_tmp_rtn] the (`Boolean`) result of `CheckError()`.


  [^1]: While you may [call a function][vba_fun_prns] without parentheses, and ignore its return value, you may _not_ thus [call a property][vba_prp_call] without assigning it ([`=`][vba_eq_op]).


## Syntax ##

These procedures have the following syntax.

```vba
Check …

CheckError([e], [type_])
```

They have the following named parameters.

| Name        | Type                           | Required | Default                     | Description                                                                                                   |
| :---------- | :----------------------------- | :------: | :-------------------------- | :------------------------------------------------------------------------------------------------------------ |
| …[^2]       | Accessor [calls][vba_prp_call] |          |                             | The calls themselves, entered as individual arguments.<br><br>This is technically a [`ParamArray`][vba_parr]. |
| `e`         | `ErrObject`[^3]                |          | [`Err`][vba_err_obj] object | The latest error, thrown during validation.                                                                   |
| `type_`[^4] | `Boolean`                      |          | `True`                      | Should `CheckError()` catch (`True`) errors for fields of the wrong type?                                     |


  [^2]: [`ParamArray`][vba_parr]s like `…` are not actually passed to a single [named argument][vba_nm_args], but rather as several nameless arguments.
  [^3]: The `ErrObject` is not a traditional "type", since there is only [one (global) instance][vba_err_typ] of the `Err` object.
  [^4]: The underscore (`_`) prevents `type_` from clashing with the [`Type`][vba_typ_kwd] keyword.


## Output ##

These procedure(s) have the following output.

  - `Check()` is inert and returns no value.  It swallows the [accessor][sob_tmp_acc] calls, unless some call throws an error on its own.
  - `CheckError()` returns `True` if no error occurred, and `False` for errors (like [type][vba_err_13]) you wish to catch (`type_ := True`).  But it ["bubbles up"][vba_ppg_err] any other error.


## Details ##

![](../med/banner_unfinished.png)


## Examples ##

[Define][vba_enum] a few [fields][sob_fld] of an SOb…

```vba
Enum Foo__Fields
	Bar
	Baz
	Qux
End Enum
```

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

Check those accessors…

```vba
Sub Check_Example(foo As Object)
	Debug.Print "Catching..."
	On Error GoTo CHECK_ERROR
	
	Debug.Print "Checking..."
	Obj_Check Foo_Bar(foo), Foo_Baz(foo), Foo_Qux(foo)
	
	Debug.Print "Succeeding..."
	
CHECK_ERROR:
	Debug.Print "Handling..."
	Debug.Print Obj_CheckError(type_ := True)
End Sub
```

…with proper fields[^5] that pass the check…

```vba
Dim foo As Object: Set foo = New_Obj("Foo")

Foo_Bar(foo) = 10
Foo_Baz(foo) = "Twenty"
Set Foo_Qux(foo) = [A1:B2]

Check_Example foo
```

> ```
> Catching...
> Checking...
> Succeeding...
> Handling...
> True
> ```

…with tampered fields that fail the check…

```vba
Obj_Field(foo, Bar) = "Ten"

Check_Example foo
```

> ```
> Catching...
> Checking...
> Handling...
> False
> ```

```vba
Set Obj_Field(foo, Bar) = New Collection

Check_Example foo
```

> ```
> Catching...
> Checking...
> Handling...
> False
> ```

```vba
Set Obj_Field(foo, Qux) = New Collection

Check_Example foo
```

> ```
> Catching...
> Checking...
> Handling...
> False
> ```

…and with irrelevant errors that ["bubble up"][vba_ppg_err].

```vba
	Debug.Print "Catching..."
	On Error GoTo CHECK_ERROR
	
	Debug.Print "Checking..."
	Dim num As Double: num = 1 / 0
	
	Debug.Print "Succeeding..."
	
CHECK_ERROR:
	Debug.Print "Handling..."
	Debug.Print Obj_CheckError(type_ := True)
```

> ```
> Catching...
> Checking...
> Handling...
> ```
> ![][sob_chk_err]


  [^5]: You may specify a [`Range`][vba_rng] with its [`.Address`][vba_rng_adr] in [shortcut notation][vba_sct_nt]: `[A1:B2]`.


## See Also ##

Topics in this project…

  - Advanced [validation][sob_tmp_chk]
  - [`Is*()`][sob_tmp_is] template
  - [`IsObj()`][sob_typo]
  - [Field accessors][sob_tmp_acc]
  - [Fields][sob_fld]
  - [Templates][sob_tmps]
  - [Setup][sob_setup] with templates

…and in VBA.

  - [Error propagation][vba_ppg_err]
  - [Error messages][vba_errs]
  - [Calling functions][vba_fun_call]
  - [Calling properties][vba_prp_call]
  - [`=`][vba_eq_op] operator
  - [`ParamArray`][vba_parr]s
  - [`Err`][vba_err_obj] object
  - [Named arguments][vba_nm_args]
  - [`Type`][vba_typ_kwd] statement
  - [`Enum`][vba_enum]erations
  - [`Range`][vba_rng]s
  - [`.Address`][vba_rng_adr] property
  - [Shortcut notation][vba_sct_nt]



  [sob_tmp_chk]:  ../src/SObTemplate.bas#L111-L140
  [sob_tmp_is]:   ../src/SObTemplate.bas#L89-L150
  [sob_typo]:     Typology.md
  [sob_tmp_acc]:  ../src/SObTemplate.bas#L171-L213
  [sob_fld]:      Field.md
  [vba_err_13]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/type-mismatch-error-13
  [vba_ppg_err]:  https://www.fastercapital.com/content/Error-Handling--Error-Handling-Excellence--Bulletproofing-Your-VBA-Code.html#Error-Bubbling-and-Propagation
  [sob_tmp_rtn]:  ../src/SObTemplate.bas#L149
  [vba_fun_prns]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures#use-parentheses-when-calling-function-procedures
  [vba_prp_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-property-procedures
  [vba_eq_op]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/equals-operator
  [vba_parr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-parameter-arrays
  [vba_err_obj]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
  [vba_err_450]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/wrong-number-of-arguments-error-450
  [vba_nm_args]:  https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures#pass-named-arguments
  [vba_err_typ]:  https://stackoverflow.com/a/55067026
  [vba_typ_kwd]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/type-statement
  [vba_enum]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_chk_err]:  ../med/vbe_error_11.png
  [vba_rng]:      https://learn.microsoft.com/office/vba/api/excel.range(object)
  [vba_rng_adr]:  https://learn.microsoft.com/office/vba/api/excel.range.address
  [vba_sct_nt]:   https://learn.microsoft.com/office/vba/excel/concepts/cells-and-ranges/refer-to-cells-by-using-shortcut-notation
  [sob_tmps]:     ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:    ../README.md#setup
  [vba_errs]:     https://learn.microsoft.com/office/vba/language/reference/error-messages
  [vba_fun_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures
