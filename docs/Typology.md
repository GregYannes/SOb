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

IsObj(x, [class])

AsObj(x, [class])
```

They have the following named parameters.

| Name    | Type                    | Required | Default | Description                               |
| :------ | :---------------------- | :------: | :------ | :---------------------------------------- |
| `obj`   | [`Collection`][vba_clx] | ✓        |         | An SOb whose "class" name you desire.     |
| `x`     | [`Variant`][vba_var]    | ✓        |         | A value which _might_ be an SOb (or not). |
| `class` | `String`                |          | `""`    | The "class" name you wish to match.       |


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

Topics in this project...

  - [`New_Obj()`][sob_cre]

...and in VBA.

  - [Properties][vba_prp]
  - [`Property Get`][vba_prp_get]
  - [`Collection`][vba_clx]s
  - [`Variant`][vba_var]s



  [vba_prp_get]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [sob_cre]:     Creation.md
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_var]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [vba_prp]:     https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
