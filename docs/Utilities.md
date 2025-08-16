# Utilities #

## Description ##

These procedures support **`SOb`** and are handy for your general use.

  - `Assign()` assigns any value (scalar or [objective][vba_isobj]) to your variable by [reference][vba_byref].
  - `Txt_Indent()` indents some text.
  - `Txt_Contains()` tests if some text contains a substring.
  - `Clx_Has()` tests if a [`Collection`][vba_clx] contains an item.
  - `Clx_Get()` safely retrieves an [`.Item`][vba_clx_itm] from a `Collection`.
  - `Clx_Set()` sets the value of an `.Item` in a `Collection`.
  - `Arr_Length()` gets the length of an [array][vba_arr].
  - `Err_Raise()` conveniently [propagates][vba_ppg_err] the latest [`Err` object][vba_err_obj].


## Syntax ##

These procedures have the following syntax.

```vba
Assign var, val

Txt_Indent(txt, [indent], [before])

Txt_Contains(txt, substring)

Clx_Has(clx, index)

Clx_Get(clx, index, [has])

Clx_Set clx, key, val

Arr_Length(arr, [dimension])

Err_Raise [e]
```

They have the following named parameters.

| Name        | Type                      | Required | Default                     | Description                                                                                                |
| :---------- | :------------------------ | :------: | :-------------------------- | :--------------------------------------------------------------------------------------------------------- |
| `var`       | [^1]                      | ✓        |                             | The [variable][vba_vrb] to which `val` should be assigned[^2].                                             |
| `val`       | [^1]                      | ✓        |                             | The value you wish to assign.                                                                              |
| `txt`       | `String`                  | ✓        |                             | Some text.                                                                                                 |
| `indent`    | `String`                  |          | [`vbTab`][vba_tab]          | The spacing used to indent `txt`.  Defaults to a standard [horizontal tab][hrz_tab] like most indentation. |
| `before`    | `Boolean`                 |          | `True`                      | Indent (`True`) the first line of `txt`?                                                                   |
| `substring` | `String`                  | ✓        |                             | The substring to seek within `txt.                                                                         |
| `clx`       | [`Collection`][vba_clx]   | ✓        |                             | Any `Collection`.                                                                                          |
| `index`     | `Long`<br><br>`String`    | ✓        |                             | The position (`Long`) or key (`String`) of the [`.Item`][vba_clx_itm] in `clx`.                            |
| `has`       | `Boolean`                 |          |                             | A flag[^2] variable to track whether `clx` actually _has_ (`True`) an `.Item` at `index`.                  |
| `key`       | `String`                  | ✓        |                             | The `key` of the `.Item` in `clx`, whose `val`ue you wish to set.                                          |
| `arr`       | [Array][vba_arr]          | ✓        |                             | Any array.                                                                                                 |
| `dimension` | `Long`                    |          | `1`                         | The [dimension][vba_arr_dmn] along which you wish to measure `arr`.                                        |
| `e`         | `ErrObject`[^3]           |          | [`Err`][vba_err_obj] object | The latest error thrown during execution.                                                                  |


  [^1]: While this is technically a [`Variant`][vba_var], it accommodates any `var`iable or `val`ue you desire.
  [^2]: The procedure updates this variable by [reference][vba_byref], which overwrites any value it originally had.
  [^3]: The `ErrObject` is not a traditional "type", since there is only [one (global) instance][vba_err_typ] of the `Err` object.


## Output ##

These procedures have the following output.

  - `Assign()` returns no value, but rather assigns the `val`ue to your `var`iable by [reference][vba_byref].
  - `Txt_Indent()` returns a `String` where each line of `txt` is indented with `indent`.
  - `Txt_Contains()` returns `True` when `txt` contains the `substring`, and `False` otherwise.
  - `Clx_Has()` returns `True` when `clx` has an [`.Item`][vba_clx_itm] at `index`, and `False` otherwise.
  - `Clx_Get()` returns a `Variant` with the value at `index`, or swallows the error when no such `.Item` exists.
  - `Clx_Set()` returns no value.  It assigns the `val`ue to any `.Item` at `key`, or [adds][vba_clx_add] it when it does not already exist.
  - `Arr_Length()` returns a `Long` with the length of `arr`, as measured along the given `dimension`.
  - `Err_Raise()` returns no value.  It [raises][vba_err_rse] the latest [`Err`][vba_err_obj]or without forcing you to specify its details.


## Details ##

![](../med/banner_unfinished.png)


## See Also ##

Topics in VBA…

  - [`IsObject()`][vba_isobj]
  - Passing [`ByRef`][vba_byref]erence
  - [`Collection`][vba_clx]s
  - [`.Item()`][vba_clx_itm] method
  - [Arrays][vba_arr]
  - [Error propagation][vba_ppg_err]
  - [`Err`][vba_err_obj] object
  - [Variables][vba_vrb]
  - [`vbTab`][vba_tab]
  - Array [dimensions][vba_arr_dmn]
  - [`Variant`][vba_var]s
  - [`.Add()`][vba_clx_add] method
  - [`.Raise()`][vba_err_rse] method

…and elsewhere.

  - [Horizontal tab][hrz_tab]



  [vba_isobj]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/isobject-function
  [vba_byref]:   https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_clx_itm]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_ppg_err]: https://www.fastercapital.com/content/Error-Handling--Error-Handling-Excellence--Bulletproofing-Your-VBA-Code.html#Error-Bubbling-and-Propagation
  [vba_err_obj]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
  [vba_vrb]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/declaring-variables
  [vba_tab]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/miscellaneous-constants
  [hrz_tab]:     https://www.ascii-code.com/9
  [vba_arr_dmn]: https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/arrays/array-dimensions
  [vba_var]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [vba_err_typ]: https://stackoverflow.com/a/55067026
  [vba_clx_add]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/add-method-visual-basic-for-applications
  [vba_err_rse]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/raise-method
