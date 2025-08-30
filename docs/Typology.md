# Typology #

## Description ##

These procedures ascertain and manipulate the "type" of an SOb.

  - `Obj_Class()` gets the "class" name of an SOb, which is a [read-only property][vba_prp_get].
  - `IsObj()` tests whether something is an SOb, originally constructed by [`New_Obj()`][sob_cre].
  - `AsObj()` ["casts"][vba_cast] something to an SOb, or dies trying.


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
| `fields` | [Array][vba_arr] of [`Enum`][vba_enum]s |          |         | All [fields][sob_fld] _required_ by an SOb of that `class`, as [enumerated][sob_doc_tmp] in your [template][sob_tmp_enm].                     |
| `strict` | `Boolean`                               |          | `True`  | Must that SOb contain _only_ (`True`) those `fields`, or may `x` contain[^1] further [`.Item`][vba_clx_itm]s[^2] (`False`) and still qualify? |


## Output ##

These procedures return the following values.

  - `Obj_Class()` returns the "class" name (`String`) of `obj`.
  - `IsObj()` returns `True` if `x` is an SOb, and `False` otherwise.  When `class` is supplied, then `IsObj()` also tests whether the "class" name matches.

> [!NOTE]
> 
> Matching for `class` is [case-insensitive][vba_cmp_mtd], much like class names in [VBA syntax][vba_naming].

  - `AsObj()` returns an SOb ([`Collection`][vba_clx]) with the original fields from `x`.  When `class` is supplied, then `AsObj()` updates the "class" name to match.

> [!WARNING]
> 
> `AsObj()` also modifies your original `x` variable, to match any `class` you supply.


## Examples ##

### Creation ###

Create an SOb of the **"Foo"** class, and examine it with `Obj_Class()`.

```vba
Dim foo As Object: Set foo = New_Obj("Foo")

Debug.Print Obj_Class(foo)
```

> ```
> Foo
> ```


### Simple Identification ###

Test whether something is an SOb…

```vba
Dim clx As Collection: Set clx = New Collection
Dim obj As Object: Set obj = New Collection

Debug.Print IsObj(foo)
Debug.Print IsObj(clx), IsObj(obj)
```

> ```
> True
> False         False
> ```

<br>

…of the **"Foo"** class or the **"Snaf"** class.

```vba
Debug.Print IsObj(foo, "Foo")
Debug.Print IsObj(foo, "foO")
Debug.Print IsObj(foo, "Snaf")
```

> ```
> True
> True
> False
> ```


### Enhanced Identification ###

Define the fields for a **"Foo"** object…

```vba
Enum Foo__Fields
	Bar
	Baz
	Qux
End Enum
```

<br>

…and test that `foo` has them all.

```vba
Dim aFields As Variant: aFields = Array(Bar, Baz, Qux)

Debug.Print "Some fields..."
Obj_Field(foo, Bar) = 10
Obj_Field(foo, Baz) = "Twenty"
Debug.Print IsObj(foo, "Foo", aFields)

Debug.Print "All fields..."
Obj_Field(foo, Qux) = 30
Debug.Print IsObj(foo, "Foo", aFields)
```

> ```
> Some fields...
> False
> All Fields...
> True
> ```

<br>

Test `foo` leniently and strictly.

```vba
Debug.Print "Exact..."
Debug.Print IsObj(foo, "Foo", aFields, strict := False)
Debug.Print IsObj(foo, "Foo", aFields, strict := True)

foo.Add "IMPOSTER"

Debug.Print "Extra..."
Debug.Print IsObj(foo, "Foo", aFields, strict := False)
Debug.Print IsObj(foo, "Foo", aFields, strict := True)
```

> ```
> Exact...
> True
> True
> Extra...
> True
> False
> ```


### Casting ###

["Cast"][vba_cast] applicable objects as SObs of the **"Snaf"** class…

```vba
Debug.Print "Declaring..."
Dim cSnaf As Collection: Set cSnaf = New Collection
Dim oSnaf As Object: Set oSnaf = New Collection

Debug.Print IsObj(cSnaf, "Snaf"), IsObj(oSnaf, "Snaf")
Debug.Print

Debug.Print "Casting..."
Set cSnaf = AsObj(cSnaf, "Snaf")
Set oSnaf = AsObj(oSnaf, "Snaf")

Debug.Print IsObj(cSnaf, "Snaf"), IsObj(oSnaf, "Snaf")
Debug.Print Obj_Class(cSnaf), Obj_Class(oSnaf)
```

> ```
> Declaring...
> False         False
> 
> Casting...
> True          True
> Snaf          Snaf
> ```

<br>

…and do likewise for `foo` which was originally a **"Foo"**.

```vba
Set foo = AsObj(foo, "Snaf")

Debug.Print IsObj(foo, "Snaf")
Debug.Print Obj_Class(foo)
```

> ```
> True
> Snaf
> ```


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
  - [Casting][vba_cast]
  - [`Collection`][vba_clx]s
  - [`Variant`][vba_var]s
  - [Arrays][vba_arr]
  - [`Enum`][vba_enum]erations
  - [`.Item()`][vba_clx_itm] method
  - [Case sensitivity][vba_cmp_mtd]
  - [Naming syntax][vba_naming]



  [^1]: As determined by [`Obj_HasFields()`][sob_flds].
  [^2]: As counted by [`Obj_FieldCount()`][sob_flds].



  [vba_prp_get]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [sob_cre]:     Creation.md
  [vba_cast]:    https://learn.microsoft.com/dotnet/visual-basic/language-reference/operators/directcast-operator
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_var]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_enum]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_fld]:     Field.md
  [sob_doc_tmp]: Setup.md#template
  [sob_tmp_enm]: ../src/SObTemplate.bas#L26-L29
  [vba_clx_itm]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
  [vba_cmp_mtd]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/strcomp-function#settings
  [vba_naming]:  https://learn.microsoft.com/office/vba/language/concepts/getting-started/visual-basic-naming-rules
  [sob_tmps]:    ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:   Setup.md
  [sob_flds]:    Fields.md
  [vba_prp]:     https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
