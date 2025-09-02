# Field Access #

## Description ##

These procedures access a field within an SOb.

  - `Obj_Field()` reads and writes the field as a [property][vba_prp].
  - `Obj_Get()` copies the field value into your variable, with protection against nonexistence.  Your ([`Property Get`][vba_prp_get]) field [accessors][sob_tmp_acc] should simply wrap `Obj_Get()`.


## Syntax ##

These procedure(s) have the following syntax.

```vba
Obj_Field(obj, field)
Obj_Field(obj, field) = val
Set Obj_Field(obj, field) = val

Obj_Get var, obj, field
```

They have the following named parameters.

| Name    | Type                    | Required | Default | Description                                                                                                |
| :------ | :---------------------- | :------: | :------ | :--------------------------------------------------------------------------------------------------------- |
| `obj`   | [`Collection`][vba_clx] | ✓        |         | An SOb whose field you wish to access.                                                                     |
| `field` | [`Enum`][vba_enum]      | ✓        |         | The field itself, as [enumerated][sob_doc_tmp] in your [template][sob_tmp_enm].                            |
| `val`   | [`Variant`][vba_var]    | ✓        |         | The value you wish to assign your field.<br><br>Use [`Set`][vba_set] when `val` is an [object][vba_isobj]. |
| `var`   | `Variant`               | ✓        |         | The variable into which `Obj_Get()` should copy the field value ([by reference][vba_byref]).               |


## Output ##

These procedures have the following output.

  - `Obj_Field()` returns a [`Variant`][vba_var] with the field value, specifically when [reading][vba_prp_get] the field.
  - `Obj_Get()` returns no value.  It copies any field value into `var` [by reference][vba_byref], but it leaves `var` untouched when no such `field` exists.


## Examples ##

### Backend Access ###

[Define][vba_enum] the `Bar` field of an SOb…

```vba
Enum Foo__Fields
	Bar
End Enum
```

<br>

…and manipulate that field.

```vba
Debug.Print "Creating..."
Dim foo1 As Object: Set foo1 = New_Obj("Foo")
Debug.Print Obj_Field(foo1, Bar)

Debug.Print "Initializing..."
Obj_Field(foo1, Bar) = 42
Debug.Print Obj_Field(foo1, Bar)
```

> ```
> Creating...
> 
> Initializing...
>  42 
> ```


### Implement Accessors ###

Implement a stable [accessor][sob_tmp_acc] for `Bar`…

```vba
Property Get Foo_Bar(foo As Object) As Integer
	Obj_Get Foo_Bar, foo, Bar
End Property

Property Let Foo_Bar(foo As Object, val As Integer)
	Let Obj_Field(foo, Bar) = val
End Property
```

<br>

…which elegantly manipulates this field…

```vba
Foo_Bar(foo1) = -1
Debug.Print Foo_Bar(foo1)
```

> ```
> -1 
> ```

<br>

…and enforces the proper type…

```vba
Foo_Bar(foo1) = "Forty-two"
```

> ![][sob_acc_err]

<br>

…but defaults to an [unitialized value][vba_emp] when the data is missing.

```vba"
Dim foo2 As Object: Set foo2 = New_Obj("Foo")
Debug.Print Foo_Bar(foo2)
```

> ```
>  0 
> ```


## See Also ##

Topics in this project…

  - [Field metadata][sob_flds]
  - [Templates][sob_tmps]
  - [Setup][sob_setup] with templates
  - [Enumerated fields][sob_tmp_enm]
  - [Field accessors][sob_tmp_acc]

…and in VBA.

  - [Properties][vba_prp]
  - [`Property Get`][vba_prp_get]
  - [`Collection`][vba_clx]s
  - [`Enum`][vba_enum]erations
  - [`Variant`][vba_var]s
  - [`Set`][vba_set] Statement
  - [`IsObject()`][vba_isobj]
  - Passing [`ByRef`erence][vba_byref]
  - [Uninitialized values][vba_emp]
  - [Error messages][vba_errs]



  [vba_prp]:     https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
  [vba_prp_get]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [sob_tmp_acc]: ../src/SObTemplate.bas#L171-L213
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_enum]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_doc_tmp]: Setup.md#template
  [sob_tmp_enm]: ../src/SObTemplate.bas#L26-L29
  [vba_var]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [vba_set]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/set-statement
  [vba_isobj]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/isobject-function
  [vba_byref]:   https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [sob_acc_err]: ../med/vbe_error_13.png
  [vba_emp]:     https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#empty
  [sob_flds]:    Fields.md
  [sob_tmps]:    ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:   Setup.md
  [vba_errs]:    https://learn.microsoft.com/office/vba/language/reference/error-messages
