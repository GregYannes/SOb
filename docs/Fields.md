# Field Metadata #

## Description ##

These functions describe [fields][sob_fld] in an SOb.

  - `Obj_FieldCount()` counts (`Long`) the fields in an SOb.
  - `Obj_HasField()` tests if an SOb has a certain field.
  - `Obj_HasFields()` tests if an SOb has a set of fields, which you supply in a single [array][vba_arr].
  - [`Obj_HasFields0()`][sob_fn0] is a convenient form of `Obj_HasFields()`, where you supply the fields directly.


## Syntax ##

These procedure(s) have the following syntax.

```vba
Obj_FieldCount(obj)

Obj_HasField(obj, field)

Obj_HasFields(obj, fields)

Obj_HasFields0(obj, …)
```

They have the following named[^1] parameters.

| Name     | Type                        | Required | Default | Description                                                                                                    |
| :------- | :-------------------------- | :------: | :------ | :------------------------------------------------------------------------------------------------------------- |
| `obj`    | [`Collection`][vba_clx]     | ✓        |         | An SOb whose field(s) you wish to assess.                                                                      |
| `field`  | [`Enum`][vba_enum]          | ✓        |         | The [field][sob_fld] itself, as [enumerated][sob_doc_tmp] in your [template][sob_tmp_enm].                     |
| `fields` | [Array][vba_arr] of `Enum`s | ✓        |         | An array of such `field`s.<br><br>This is best achieved via [`Array()`][vba_arr_fn].                           |
| …[^1]    | `Enum`s                     |          |         | The fields themselves, entered as individual arguments.<br><br>This is technically a [`ParamArray`][vba_parr]. |


## Output ##

These procedure(s) have the following output.

  - `Obj_FieldCount()` returns the number (`Long`) of [fields][sob_fld] in an SOb.  This _excludes_ metadata like its ["class" name][sob_typo].

> [!NOTE]
> 
> This is technically an _upper bound_ on the field count, since an SOb is [built on][sob_doc_clx] a [`Collection`][vba_clx] which may hold extra [`.Item`][vba_clx_itm]s beyond its fields.

  - `Obj_HasField()` returns `True` if `field` is present in `obj`, and `False` otherwise.
  - `Obj_HasFields()` returns `True` if _all_ `fields` are present in `obj`, and `False` otherwise.
  - `Obj_HasFields0()` does likewise for all fields in `…`.


## Examples ##

### Count Fields ###

[Define][vba_enum] several [fields][sob_fld] of an SOb…

```vba
Enum Foo__Fields
	Bar
	Baz
	Qux
End Enum
```

<br>

…and count them.

```vba
Debug.Print "Creating..."
Dim foo As Object: Set foo = New_Obj("Foo")
Debug.Print Obj_FieldCount(foo)

Debug.Print "Initializing..."
Obj_Field(foo, Bar) = 10
Obj_Field(foo, Baz) = "Twenty"
Obj_Field(foo, Qux) = 30
Debug.Print Obj_FieldCount(foo)
```

> ```
> Creating...
>  0 
> Initializing...
>  3 
> ```

<br>

"Trick" the `Obj_FieldCount()` by manipulating this SOb as a [`Collection`][vba_clx]: namely [removing][vba_clx_rmv] its [final][vba_clx_cnt] field (`Qux`) and [adding][vba_clx_add] a "dummy" [`.Item`][vba_clx_itm] in its place.

```vba
Debug.Print "Removing..."
foo.Remove foo.Count
Debug.Print Obj_Field(foo, Qux)
Debug.Print Obj_FieldCount(foo)

Debug.Print "Adding..."
foo.Add "IMPOSTER"
Debug.Print Obj_FieldCount(foo)
```

> ```
> Removing...
> 
>  2 
> Adding...
>  3 
> ```


### Check Existence ###

Test whether each field still exists.

```vba
Debug.Print Obj_HasField(foo, Bar)
Debug.Print Obj_HasField(foo, Baz)
Debug.Print Obj_HasField(foo, Qux)
```

> ```
> True
> True
> False
> ```

<br>

Test programmatically whether _all_ fields exist…

```vba
Dim f1 As Variant: f1 = Array(Bar, Baz)
Debug.Print Obj_HasFields(foo, f1)

Dim f2 As Variant: f2 = Array(Bar, Baz, Qux)
Debug.Print Obj_HasFields(foo, f2)
```

> ```
> True
> False
> ```

<br>

…and test them manually.

```vba
Debug.Print Obj_HasFields0(foo, Bar, Baz)
Debug.Print Obj_HasFields0(foo, Bar, Baz, Qux)
```

> ```
> True
> False
> ```


## See Also ##

Topics in this project…

  - [Fields][sob_fld]
  - [`*0()`][sob_fn0] family
  - [Templates][sob_tmps]
  - [Setup][sob_setup] with templates
  - [Enumerated fields][sob_tmp_enm]
  - [Field accessors][sob_tmp_acc]

…and in VBA.

  - [Arrays][vba_arr]
  - [`Collection`][vba_clx]s
  - [`Enum`][vba_enum]erations
  - [`ParamArray`][vba_parr]s
  - [Named arguments][vba_nm_args]
  - [`.Item()`][vba_clx_itm] method
  - [`.Remove()`][vba_clx_rmv] method
  - [`.Add()`][vba_clx_add] method



  [^1]: [`ParamArray`][vba_parr]s like `…` are not actually passed to a single [named argument][vba_nm_args], but rather as several nameless arguments.



  [sob_fld]:     Field.md
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [sob_fn0]:     Zero.md
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_enum]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_doc_tmp]: Setup.md#template
  [sob_tmp_enm]: ../src/SObTemplate.bas#L26-L29
  [vba_arr_fn]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/array-function
  [vba_parr]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-parameter-arrays
  [sob_typo]:    Typology.md
  [sob_doc_clx]: Creation.md#details
  [vba_clx_itm]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
  [vba_clx_rmv]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/remove-method-visual-basic-for-applications
  [vba_clx_cnt]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/count-property-visual-basic-for-applications
  [vba_clx_add]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/add-method-visual-basic-for-applications
  [sob_tmps]:    ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:   Setup.md
  [sob_tmp_acc]: ../src/SObTemplate.bas#L171-L213
  [vba_nm_args]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
