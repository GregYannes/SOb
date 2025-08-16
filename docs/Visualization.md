# Visualization #

## Description ##

These functions support ["pretty printing"][pprint] for SObs.

  - `Obj_Format()` formats a textual representation of an SOb.
  - `Obj_Print()` also [prints][vba_print] that format to the [console][vbe_immed].
  - [`Obj_Print0()`][sob_fn0] is a primitive form of `Obj_Print()`, which prints to the console without formatting.
  - `Obj_FormatFields()` formats a textual representation of the [fields][sob_fld] in an SOb, which you supply in a [single array][vba_arr].  This representation is often used for the `details` of an SOb.
  - [`Obj_FormatFields0()`][sob_fn0] is a convenient form of `Obj_FormatFields()`, where you supply the fields directly.


## Syntax ##

These functions have the following syntax.

```vba
Obj_Format(obj, [depth], [plain], [pointer], [summary], [details], [preview], [indent], [orphan])

Obj_Print(obj, [depth], [plain], [pointer], [summary], [details], [preview], [indent], [orphan])

Obj_Print0([format])

Obj_FormatFields(fields, [separator])

Obj_FormatFields0(…)
```

They have the following named parameters.

| Name        | Type                              | Required | Default                  | Description                                                                                                                                                                                 |
| :---------- | :-------------------------------- | :------: | :----------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `obj`       | [`Collection`][vba_clx]           | ✓        |                          | An SOb you wish to visualize.                                                                                                                                                               |
| `depth`     | `Integer`                         |          | `1`                      | The depth to which the visualization should expand: `0` shows a `summary`, while positive expands to any `details`.                                                                         |
| `plain`     | `Boolean`                         |          | `False`                  | Format the SOb in plain format (`True`) rather than rich (`False`)?  See [**Details**][sob_vis_dtl] for appearance.                                                                         |
| `pointer`   | `Boolean`                         |          | `False`                  | Fall back to showing the [pointer][vba_ptr] (`True`) for `obj`, when we have neither `summary` nor `preview`?                                                                               |
| `summary`   | `String`                          |          | `""`                     | An expression summarizing the contents of `obj` on a single line, when `depth := 0`.  See `preview` and `pointer` for fallbacks when `summary := ""`.                                       |
| `details`   | `String`                          |          | `""`                     | Expressions detailing the contents of `obj` across multiple lines, when `depth > 0`.  See [**Details**][sob_vis_dtl] for formatting with `Obj_FormatFields*()`.                             |
| `preview`   | `Boolean`                         |          | `False`                  | Fall back to showing a preview of the `details`, when we have no `summary`?  See [**Details**][sob_vis_dtl] for appearance.                                                                 |
| `indent`    | `String`                          |          | [`vbTab`][vba_tab]       | The indentation used for nesting `details`.  Defaults to a standard [horizontal tab][hrz_tab] like most indentation.                                                                        |
| `orphan`    | `Boolean`                         |          | `True`                   | Should a single line of `details` still be nested (`True`) or remain on a single line (`False`)?                                                                                            |
| `format`    | `String`                          |          | `""`                     | Output for the console, which should already be formatted as desired.                                                                                                                       |
| `fields`    | [Array][vba_arr] of `String`s[^1] | ✓        |                          | An array with pairs of (textual) expressions: a field name followed by its value.  See [**Details**][sob_vis_dtl] for appearance.<br><br>This is best achieved via [`Array()`][vba_arr_fn]. |
| `separator` | `String`                          |          | [`vbNewLine`][vba_newln] | The textual separator displayed between each pairing and the next.  Defaults to the [system newline][sys_newln], so each pair (`.field = value`) gets its own line.                         |
| …[^2]       | `String`s                         |          |                          | The pairs themselves, entered as individual arguments.<br><br>This is technically a [`ParamArray`][vba_parr].                                                                               |


> [!NOTE]
> 
> Be sure to use the [`vbNewLine`][vba_newln] for line breaks, when you assemble (say) `details` and other such text.  This uses the newline [specific to the system][sys_newln], and ensures that `Obj_FormatFields*()` and [`Txt_Indent()`][sob_utils] work as expected.


  [^1]: You may use a `String()` array, or a `Variant()` array containing `String`s.
  [^2]: [`ParamArray`][vba_parr]s like `…` are not actually passed to a single [named argument][vba_nm_args], but rather as several nameless arguments.


## Output ##

These functions have the following output.

  - `Obj_Format()` returns a `String` with the formatted representation of `obj`.
  - `Obj_Print()` returns the same `String` and also [prints][vba_print] it to the [console][vbe_immed].
  - `Obj_Print0()` returns `format` and also prints it to the console "as is".
  - `Obj_FormatFields()` returns a `String` with the formatted representation (`.field = value`) of all pairs in `fields`.
  - `Obj_FormatFields0()` does likewise for all pairs in `…`.


## Details ##

![](../med/banner_unfinished.png)


## See Also ##

Topics in this project…

  - [`*0()`][sob_fn0] family
  - [Fields][sob_fld]

…in VBA…

  - [`.Print()`][vba_print] method
  - [Immediate window][vbe_immed] in the [Visual Basic Editor][vbe] (VBE)
  - [Arrays][vba_arr]
  - [`Collection`][vba_clx]s
  - [Pointers][vba_ptr]
  - [`vbTab`][vba_tab]
  - [`Array()`][vba_arr_fn]
  - [`vbNewLine`][vba_newln]
  - [`ParamArray`][vba_parr]s
  - [Named arguments][vba_nm_args]

…and elsewhere.

  - [Pretty-printing][pprint]
  - [Horizontal tab][hrz_tab]
  - [System newline][sys_newln]



  [pprint]:      https://en.wikipedia.org/wiki/Pretty-printing
  [vba_print]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/print-method
  [vbe_immed]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/immediate-window
  [sob_fn0]:     Zero.md
  [sob_fld]:     Field.md
  [vba_arr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_clx]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [sob_vis_dtl]: #details
  [vba_ptr]:     https://classicvb.net/tips/varptr
  [vba_tab]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/miscellaneous-constants
  [hrz_tab]:     https://www.ascii-code.com/9
  [vba_arr_fn]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/array-function
  [vba_newln]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/miscellaneous-constants
  [sys_newln]:   https://learn.microsoft.com/dotnet/api/system.environment.newline?view=net-9.0#property-value
  [vba_parr]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-parameter-arrays
  [sob_utils]:   Utilities.md
  [vba_nm_args]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
  [vbe]:         https://learn.microsoft.com/office/vba/library-reference/concepts/getting-started-with-vba-in-office#macros-and-the-visual-basic-editor
