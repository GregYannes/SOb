# Zero Suffix #

## Description ##

The "`0`" suffix denotes a convenience function, which processes its arguments _"as is"_.  In the spirit of [`paste()`][r_paste] and [`paste0()`][r_paste] in [R][r_lang], each `*0()` function is a variation on another `*()` function, which strips away nonessential "settings" in favor of handy defaults.

Currently **`SOb`** offers three such functions:

  1. [`Obj_HasFields0()`][sob_flds]
  1. [`Obj_Print0()`][sob_vis]
  1. [`Obj_FormatFields0()`][sob_vis]


## Syntax ##

When illustrating syntax, an ellipsis (`…`) denotes a [`ParamArray`][vba_parr] of arguments.  You enter such arguments individually, rather than (say) passing them in a single [array][vba_arr].

```vba
Obj_HasFields0(obj, …)

Obj_FormatFields0(…)
```


## Details ##

![](../med/banner_unfinished.png)


## See Also ##

Topics in this project...

  - [`Obj_HasFields()`][sob_flds]
  - [`Obj_Print()`][sob_vis]
  - [`Obj_FormatFields`][sob_vis]

...and elsewhere.

  - [R][r_lang] language
  - [`paste()`][r_paste] and [`paste0()`][r_paste]



  [r_paste]:  https://rdrr.io/r/base/paste.html
  [r_lang]:   https://www.r-project.org/about.html
  [sob_flds]: Fields.md
  [sob_vis]:  Visualization.md
  [vba_parr]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-parameter-arrays
  [vba_arr]:  https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
