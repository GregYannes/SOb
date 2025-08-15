# Field Metadata #

## Description ##

These functions provide metadata about [fields][sob_fld] in an SOb.

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
Obj_HasFields(obj, …)
```

They have the following named[^1] parameters.

| Name     | Type                    | Required | Default | Description                                                                                                    |
| :------- | :---------------------- | :------: | :------ | :------------------------------------------------------------------------------------------------------------- |
| `obj`    | [`Collection`][vba_clx] | ✓        |         | An SOb whose field(s) you wish to assess.                                                                      |
| `field`  | [`Enum`][vba_enum]      | ✓        |         | The [field][sob_fld] itself, as [enumerated][sob_rdm_tmp] in your [template][sob_tmp_enm].                     |
| `fields` | Array of `Enum`s        | ✓        |         | An [array][vba_arr] of such `field`s.                                                                          |
| …[^1]    | `Enum`s                 |          |         | The fields themselves, entered as individual arguments.<br><br>This is technically a [`ParamArray`][vba_parr]. |


  [^1]: [`ParamArray`][vba_parr]s like `…` are not actually passed to a single [named argument][vba_nm_args], but rather as several nameless arguments.


## Output ##

These procedure(s) have the following output.

  - `Obj_FieldCount()` returns the number (`Long`) of [fields][sob_fld] in an SOb.  This _excludes_ metadata like its ["class" name][sob_typo].

> [!NOTE]
> 
> This is technically an _upper bound_ on the field count, since an SOb is [built on][sob_rdm_clx] a [`Collection`][vba_clx] which may hold extra [`.Item`][vba_clx_itm]s beyond its fields.

  - `Obj_HasField()` returns `True` if `field` is present in `obj`, and `False` otherwise.
  - `Obj_HasFields()` returns `True` if _all_ `fields` are present in `obj`, and `False` otherwise.
  - `Obj_HasFields0()` does likewise for all fields in `…`.


## Details ##

![](../med/banner_unfinished.png)


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
  - [`ParamArray`s][vba_parr]
  - [`.Item()`][vba_clx_itm] method



  [sob_fld]:     Field.md
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [sob_fn0]:     Zero.md
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_enum]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_rdm_tmp]: ../README.md#template
  [sob_tmp_enm]: ../src/SObTemplate.bas#L26-L29
  [vba_parr]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-parameter-arrays
  [vba_nm_args]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
  [sob_typo]:    Typology.md
  [sob_rdm_clx]: ../README.md#new-solution
  [vba_clx_itm]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
  [sob_tmps]:    ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:    ../README.md#setup
  [sob_tmp_acc]: ../src/SObTemplate.bas#L171-L213
