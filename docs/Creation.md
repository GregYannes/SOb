# Creation #

## Description ##

These procedures prepare a new SOb for use.

  - `New_Obj()` constructs a new SOb from scratch.  This mimics the [`New`][vba_new] operator in VBA.
  - `Obj_Initialize()` initializes an SOb.


## Syntax ##

These procedure have the following syntax.

```vba
New_Obj(class)

Obj_Initialize obj, class
```

They have the following named parameters.

| Name    | Type                    | Required | Default | Description                                                    |
| :------ | :---------------------- | :------: | :------ | :------------------------------------------------------------- |
| `obj`   | [`Collection`][vba_clx] | ✓        |         | An SOb you wish to initialize.                                 |
| `class` | `String`                | ✓        |         | The "class" name of your SOb.  See [**Details**][sob_cre_dtl]. |


## Output ##

These procedures have the following output.

  - `New_Obj()` returns a [`New Collection`][vba_new_clx], with `class` as its ["class" property][sob_typo].
  - `Obj_Initialize()` operates on `obj` but returns no value.  It ensures that `obj` is an initialized `Collection`, with `class` as its "class" property.


## Details ##

![](../med/banner_unfinished.png)


## See Also ##

Topics in this project...

  - [`Obj_Class()`][sob_typo]

...and in VBA.

  - [`New`][vba_new] operator
  - [`Collection`][vba_clx]s



  [vba_new]:     https://learn.microsoft.com/dotnet/visual-basic/language-reference/operators/new-operator
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [sob_cre_dtl]: #details
  [vba_new_clx]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object#remarks
  [sob_typo]:    Typology.md
