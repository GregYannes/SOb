# Typology #

## Description ##

These procedures ascertain and manipulate the "type" of an SOb.

  - `Obj_Class()` gets the "class" name of an SOb, which is a [read-only property][vba_prp_get].
  - `IsObj()` tests whether something is an SOb, originally constructed by [`New_Obj()`][sob_cre].
  - `AsObj()` casts something to an SOb, or dies trying.


## Syntax ##

These procedures have the following syntax.

```vba
Obj_Class(obj)

IsObj(x, [class], [fields], [strict])

AsObj(x, [class])
```

They have the following named parameters.

| Name     | Type                                    | Required | Default | Description                                                                                                                                   |
| :------- | :-------------------------------------- | :------: | :------ | :-------------------------------------------------------------------------------------------------------------------------------------------- |
| `obj`    | [`Collection`][vba_clx]                 | ✓        |         | An SOb whose "class" name you desire.                                                                                                         |
| `x`      | [`Variant`][vba_var]                    | ✓        |         | A value which _might_ be an SOb (or not).                                                                                                     |
| `class`  | `String`                                |          |         | The "class" name you wish to match.                                                                                                           |
| `fields` | [Array][vba_arr] of [`Enum`][vba_enum]s |          |         | All [fields][sob_fld] _required_ by an SOb of that `class`, as [enumerated][sob_rdm_tmp] in your [template][sob_tmp_enm].                     |
| `strict` | `Boolean`                               |          | `True`  | Must that SOb contain _only_ (`True`) those `fields`, or may `x` contain[^1] further [`.Item`][vba_clx_itm]s[^2] (`False`) and still qualify? |


  [^1]: As determined by [`Obj_HasFields()`][sob_flds].
  [^2]: As counted by [`Obj_FieldCount()`][sob_flds].


## Output ##

These procedures return the following values.

  - `Obj_Class()` returns the "class" name (`String`) of `obj`.
  - `IsObj()` returns `True` if `x` is an SOb, and `False` otherwise.  When `class` is supplied, then `IsObj()` also tests whether the "class" name matches.
  - `AsObj()` returns an SOb ([`Collection`][vba_clx]) with the original fields from `x`.  When `class` is supplied, then `AsObj()` updates the "class" name to match.

> [!WARNING]
> 
> `AsObj()` also modifies `x` itself, to match any `class` you supply.


## Details ##

![](../med/banner_unfinished.png)


## See Also ##

Topics in this project…

  - [`New_Obj()`][sob_cre]
  - [Fields][sob_fld]
  - [Templates][sob_tmps]
  - [Setup][sob_setup] with templates
  - [Enumerated fields][sob_tmp_enm]
  - [`Obj_HasFields()`][sob_flds]
  - [`Obj_CountFields()`][sob_flds]

…and in VBA.

  - [Properties][vba_prp]
  - [`Property Get`][vba_prp_get]
  - [`Collection`][vba_clx]s
  - [`Variant`][vba_var]s
  - [Arrays][vba_arr]
  - [`Enum`][vba_enum]erations
  - [`.Item()`][vba_clx_itm] method



  [vba_prp_get]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [sob_cre]:     Creation.md
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_var]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_enum]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_fld]:     Field.md
  [sob_rdm_tmp]: ../README.md#template
  [sob_tmp_enm]: ../src/SObTemplate.bas#L26-L29
  [vba_clx_itm]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
  [sob_flds]:    Fields.md
  [sob_tmps]:    ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:   ../README.md#setup
  [vba_prp]:     https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
