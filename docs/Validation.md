# Validation #

## Description ##

These procedures help you validate [advanced implementations][sob_tmp_chk] of [`Is*()`][sob_tmp_is], where invalid input would otherwise crash.  Such validation helps your `Is*()` function identify your SObs precisely, beyond the basics supported by [`IsObj()`][sob_typo].

  - `Check()` conveniently calls your [accessors][sob_tmp_acc], merely to check that they work.  This lets you ignore the [field][sob_fld] value[^1] itself.
  - `CheckError()` catches certain errors (like [type mismatch][vba_err_13]) that invalidate the check, but it [propagates][vba_ppg_err] all other errors.  Your `Is*` function should [return][sob_tmp_rtn] the (`Boolean`) result of `CheckError()`.


  [^1]: While you may [call a function][vba_fun_prns] without parentheses, and ignore its return value, you may _not_ thus [call a property][vba_prp_call] without assigning it ([`=`][vba_eq_op]).


## Syntax ##

These procedures have the following syntax.

```vba
Check …

CheckError([e], [type_])
```

They have the following named parameters.

| Name        | Type                           | Required | Default                     | Description                                                                                                                                                                             |
| :---------- | :----------------------------- | :------: | :-------------------------- | :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| …[^2]       | Accessor [calls][vba_prp_call] |          |                             | The calls themselves, entered as individual arguments.<br><br>This is technically a [`ParamArray`][vba_parr].                                                                           |
| `e`         | `ErrObject`[^3]                |          | [`Err`][vba_err_obj] object | The latest error, thrown during validation.                                                                                                                                             |
| `type_`[^4] | `Boolean`                      |          | `True`                      | Should `CheckError()` catch (`True`) errors for fields of the wrong type?<ul><li>[**`Error 13`**][vba_err_13] for scalars.</li><li>[**`Error 450`**][vba_err_450] for objects.</li><ul> |


  [^2]: [`ParamArray`][vba_parr]s like `…` are not actually passed to a single [named argument][vba_nm_args], but rather as several nameless arguments.
  [^3]: The `ErrObject` is not a traditional "type", since there is only [one (global) instance][vba_err_typ] of the `Err` object.
  [^4]: The underscore (`_`) prevents `type_` from clashing with the [`Type`][vba_typ_kwd] keyword.


## Output ##

These procedure(s) have the following output.

  - `Check()` is inert and returns no value.  It swallows the [accessor][sob_tmp_acc] calls, unless some call throws an error on its own.
  - `CheckError()` returns `True` if no error occurred, and `False` for errors (like [type][vba_err_13]) you wish to catch (`type_ := True`).  But it ["bubbles up"][vba_ppg_err] any other error.


## Details ##

![](../med/banner_unfinished.png)


## See Also ##

Topics in this project…

  - Advanced [validation][sob_tmp_chk]
  - [`Is*()`][sob_tmp_is] template
  - [`IsObj()`][sob_typo]
  - [Field accessors][sob_tmp_acc]
  - [Fields][sob_fld]
  - [Templates][sob_tmps]
  - [Setup][sob_setup] with templates

…and in VBA.

  - [Error propagation][vba_ppg_err]
  - [Error messages][vba_errs]
  - [Calling functions][vba_fun_call]
  - [Calling properties][vba_prp_call]
  - [`=`][vba_eq_op] operator
  - [`ParamArray`][vba_parr]s
  - [`Err`][vba_err_obj] object
  - [Named arguments][vba_nm_args]
  - [`Type`][vba_typ_kwd] statement



  [sob_tmp_chk]:  ../src/SObTemplate.bas#L111-L140
  [sob_tmp_is]:   ../src/SObTemplate.bas#L89-L150
  [sob_typo]:     Typology.md
  [sob_tmp_acc]:  ../src/SObTemplate.bas#L171-L213
  [sob_fld]:      Field.md
  [vba_err_13]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/type-mismatch-error-13
  [vba_ppg_err]:  https://www.fastercapital.com/content/Error-Handling--Error-Handling-Excellence--Bulletproofing-Your-VBA-Code.html#Error-Bubbling-and-Propagation
  [sob_tmp_rtn]:  ../src/SObTemplate.bas#L149
  [vba_fun_prns]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures#use-parentheses-when-calling-function-procedures
  [vba_prp_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-property-procedures
  [vba_eq_op]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/equals-operator
  [vba_parr]:     https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-parameter-arrays
  [vba_err_obj]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
  [vba_err_450]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/wrong-number-of-arguments-error-450
  [vba_nm_args]:  https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
  [vba_err_typ]:  https://stackoverflow.com/a/55067026
  [vba_typ_kwd]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/type-statement
  [sob_tmps]:     ../../../search?type=code&q=path:src/*Template.bas
  [sob_setup]:    ../README.md#setup
  [vba_errs]:     https://learn.microsoft.com/office/vba/language/reference/error-messages
  [vba_fun_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures
