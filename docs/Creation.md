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
| `class` | `String`                | ✓        |         | The "class" name of your SOb.  See [**Examples**][sob_cre_ex]. |


## Output ##

These procedures have the following output.

  - `New_Obj()` returns a [`New Collection`][vba_new_clx], with `class` as its ["class" property][sob_typo].
  - `Obj_Initialize()` operates on `obj` but returns no value.  It ensures that `obj` is an initialized `Collection` with a "class" property.


## Examples ##

### Creation ###

Create an SOb of the **"Foo"** class, and examine it with [`Obj_Class()`][sob_typo].

```vba
Dim foo As Object: Set foo = New_Obj("Foo")

Debug.Print Obj_Class(foo)
```

> ```
> Foo
> ```


### Initialization ###

Ensure applicable objects are initialized as SObs of the **"Snaf"** class…

```vba
Debug.Print "Declaring..."
Dim cSnaf As Collection, oSnaf As Object

Debug.Print cSnaf Is Nothing, oSnaf Is Nothing
Debug.Print

Debug.Print "Initializing..."
Obj_Initialize cSnaf, "Snaf"
Obj_Initialize oSnaf, "Snaf"

Debug.Print cSnaf Is Nothing, oSnaf Is Nothing
Debug.Print Obj_Class(cSnaf), Obj_Class(oSnaf)
```

> ```
> Declaring...
> True          True
> 
> Initializing...
> False         False
> Snaf          Snaf
> ```

<br>

…but leave `foo` untouched because it is _already_ a **"Foo"**.

```vba
Obj_Initialize foo, "Snaf"

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
  [sob_cre_ex]:  #examples
  [vba_new_clx]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object#remarks
  [sob_typo]:    Typology.md
  [vba_obj]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/object-data-type
