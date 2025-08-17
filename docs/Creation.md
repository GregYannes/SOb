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
  - `Obj_Initialize()` operates on `obj` but returns no value.  It ensures that `obj` is an initialized `Collection` with a "class" property.


## Details ##

![](../med/banner_unfinished.png)


## Examples ##

Create an SOb of the **"Foo"** class, and examine it with [`Obj_Class()`][sob_typo].

```vba
Dim foo As Object: Set foo = New_Obj("Foo")

Debug.Print Obj_Class(foo)
```

> ```
> Foo
> ```

<br>

Ensure applicable objects are initialized as SObs of the **"Bar"** class…

```vba
Debug.Print "Declaring..."
Dim cBar As Collection, oBar As Object

Debug.Print cBar Is Nothing, oBar Is Nothing
Debug.Print

Debug.Print "Initializing..."
Obj_Initialize cBar, "Bar"
Obj_Initialize oBar, "Bar"

Debug.Print cBar Is Nothing, oBar Is Nothing
Debug.Print Obj_Class(cBar), Obj_Class(oBar)
```

> ```
> Declaring...
> True          True
> 
> Initializing...
> False         False
> Bar           Bar
> ```

…but fail to do likewise for `foo` which is _already_ a **"Foo"**.

```vba
Obj_Initialize foo, "Bar"

Debug.Print Obj_Class(foo)
```

> ```
> Foo
> ```


## See Also ##

Topics in this project…

  - [`Obj_Class()`][sob_typo]

…and in VBA.

  - [`New`][vba_new] operator
  - [`Collection`][vba_clx]s
  - [`Object`][vba_obj]s



  [vba_new]:     https://learn.microsoft.com/dotnet/visual-basic/language-reference/operators/new-operator
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [sob_cre_dtl]: #details
  [vba_new_clx]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object#remarks
  [sob_typo]:    Typology.md
  [vba_obj]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/object-data-type
