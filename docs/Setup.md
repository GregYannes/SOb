# Setup #

Setup is quick and painless with [handy templates][sob_tmps].  Simply fill out the [`TODO`][sob_todos]s and paste the result in your module!


## Consolidated ##

To consolidate everything within your existing module, fill out [`SnippetTemplate.bas`][sob_snp_tmp] and paste into your module.  Then paste [`Snippet.bas`][sob_snp] alongside it.


## Outsourced ##

To outsource the **`SOb`** framework to a single external dependency, fill out [`SObTemplate.bas`][sob_mod_tmp] and paste into your module.  Then instruct your end user to import the [`SOb.bas`][sob_mod] module, which you may reference as a [submodule][ghub_submod] in your repo.


> [!WARNING]
> 
> By outsourcing, you reduce your security against tampering!  The ["encryption"][sob_secure] is no longer scoped to your module, so others can overwrite[^1] your "private" fields.


  [^1]: Others cannot typically _overwrite_ the value of a "private" field in your SOb—though they can _remove_ the field, which effectively resets it to an uninitialized state.
    
    However, if you [outsource the framework][sob_outsrc] from your module to the [**`SOb`** module][sob_mod], then others _can_ overwrite it via [`SOb.Obj_Field()`][sob_fld].


## Template ##

Fill out either template according to these steps:

  1. [`TODO`][sob_todo_1]: Replace every [`*`][sob_tmp_ast] with the (["class"][sob_cls]) name you desire for your SOb.  A simple **Find & Replace** should suffice.
     
     So if you name your SOb something like "`Foo`", this should yield a `FOO_CLS` [constant][vba_const] atop your module; along with a `Foo__Field` [enumeration][vba_enum] below; followed by [procedures][vba_proc] of the form `Foo_…()` with a `foo` [argument][vba_arg].
     
  1. [`TODO`][sob_todo_2]: Enumerate the [fields][sob_tmp_enm] you desire for your SOb.  Currently there are three placeholders for these fields: [`FieldOne`][sob_tmp_f1] and [`FieldTwo`][sob_tmp_f2] and [`FieldThree`][sob_tmp_f3].  Feel free to **Find & Replace** these, and to append further fields of your own.
     
     This way, you can specify a field to (say) [`Obj_Field()`][sob_fld] by simply selecting it from the [dropdown][vbe_drop] for [`Foo__Field.…`][sob_tmp_fld].
     
  1. [`TODO`][sob_todo_9]: Implement [accessors][sob_tmp_acc] for your [`Foo__Field.…`][sob_tmp_fld] fields.  Each should be a [`Property`][vba_prp] of the form `Foo_…(ByRef foo As Object)`, and you may restrict it to internal usage via the [`Private`][vba_priv] statement.
     
     Simply wrap [`Obj_Get()`][sob_fld] with a [`Property Get`][vba_prp_get] to [_retrieve_][sob_tmp_get] a field…
     
     ```vba
     Property Get Foo_FieldOne(ByRef foo As Object) As Integer
     	Obj_Get Foo_FieldOne, foo, Foo__Field.FieldOne
     End Property
     ```
     
     …but implement a [`Property Let`][vba_prp_let] to _update_ a [scalar field][sob_tmp_scl]…
     
     ```vba
     Property Let Foo_FieldOne(ByRef foo As Object, ByVal val As Integer)
     	Let Foo_FieldOne = val
     End Property
     ```
     
     …or a [`Property Set`][vba_prp_set] to update an [objective field][sob_tmp_obj].
     
     ```vba
     Property Set Foo_FieldTwo(ByRef foo As Object, ByRef val As Range)
     	Set Foo_FieldTwo = val
     End Property
     ```
     
  1. [`TODO`][sob_todo_3]: Initialize the values for your fields, within [`Foo_Initialize()`][sob_tmp_ini].  Use [`Obj_HasField()`][sob_flds] to test whether a field exists, and when it does not, use your [accessor][sob_tmp_acc] to set its initial value.
     
  1. [`TODO`][sob_todo_4]: List all your [`Foo__Field.…`][sob_tmp_fld] fields in the [`Array(…)`][sob_tmp_arr] call, within [`IsFoo()`][sob_tmp_is].
     
     This way, `IsFoo()` checks that a "Foo" object has all its fields.
     
  1. [`TODO`][sob_todo_8]: Using your accessors, [assign each field][sob_tmp_asn] from `obj` to its corresponding field in `AsFoo`, within [`AsFoo()`][sob_tmp_as].
     
     This way, `AsFoo()` can cast any input (`x`) to a "Foo" object, by extracting fields from the former into the latter.
     
  1. [`TODO`][sob_todo_11]: Create any `summary` or `details` you wish, to visually represent your object within [`Foo_Format()`][sob_tmp_fmt].  The **`SOb`** framework _automatically_ formats these for you: summaries display on a single line…
     
     > ```
     > <Foo[…]>
     > ```
     
     …and details display across multiple lines:
     
     > ```
     > <Foo: {
     > 	…
     > 	…
     > 	…
     > }>
     > ```
     
  1. [`TODO`][sob_todo_12]: Pass any `summary` or `details` to [`Obj_Format()`][sob_vis], along with all arguments from [`Foo_Format()`][sob_tmp_fmt].
     
     This way, others can apply various settings when printing your "Foo" object, including developers who wish to build their own SObs upon "Foo".

You may _optionally_ enhance "Foo" with further steps:

  9. [`TODO`][sob_todo_10]: Implement any [methods][sob_tmp_mtd] you desire, which operate on your "Foo" object.  Each should be a [`Function`][vba_fun] or [`Sub`routine][vba_sub] of the form `Foo_…(ByRef foo As Object, …)` where `foo` is followed by any [arguments][vba_arg] needed by the method.  You may restrict it to internal usage via the [`Private`][vba_priv] statement.
     
  9. [`TODO`][sob_todo_5]: Call all your field accessors like [`Foo_FieldOne()`][sob_tmp_p1], in the [`Check …`][sob_tmp_chk] call within [`IsFoo()`][sob_tmp_is].
     
     This way, `IsFoo()` also validates that the "Foo" fields are of the expected type, and so forth.
     
  9. [`TODO`][sob_todo_7]: Specify which validation errors (like type) you wish to catch, via arguments to the [`Obj_CheckError(…)`][sob_tmp_err] call within [`IsFoo()`][sob_tmp_is].  See [`Obj_CheckError()`][sob_err_args] for details.
     
     This way, `IsFoo()` returns `False` when errors disqualify the input (`x`) as a "Foo" object, while ["bubbling up"][vba_ppg_err] other errors for (say) improper usage.
     
  9. [`TODO`][sob_todo_6]: Customize any [further validation][sob_tmp_vld] you wish [`IsFoo()`][sob_tmp_is] to perform.  Each validation step should assign a `Boolean` value to `IsFoo`…
     
     ```vba
     	IsFoo = …
     ```
     
     …and finish by [short-circuiting][sob_tmp_cir] when `False`.
     
     ```vba
     	If Not IsFoo Then Exit Function
     ```

Now you are ready to work with "Foo" objects, within your module and elsewhere!



  [sob_tmps]:     ../../../search?type=code&q=path:src/*Template.bas
  [sob_todos]:    ../../../search?type=code&q=path:src/*Template.bas+content:TODO:
  [sob_snp_tmp]:  ../src/SnippetTemplate.bas
  [sob_snp]:      ../src/Snippet.bas
  [sob_mod_tmp]:  ../src/SObTemplate.bas
  [sob_mod]:      ../src/SOb.bas
  [ghub_submod]:  https://github.blog/open-source/git/working-with-submodules
  [sob_secure]:   ../src/SOb.bas#L489-L504
  [sob_outsrc]:   ../README.md#outsourced
  [sob_fld]:      Field.md
  [sob_todo_1]:   ../src/SObTemplate.bas#L6
  [sob_tmp_ast]:  ../../../search?type=code&q=path:src/*Template.bas+content:*
  [sob_cls]:      ../README.md#typology
  [vba_const]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/declaring-constants
  [vba_enum]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [vba_proc]:     https://learn.microsoft.com/office/vba/language/how-to/create-a-procedure
  [vba_arg]:      https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
  [sob_todo_2]:   ../src/SObTemplate.bas#L25
  [sob_tmp_enm]:  ../src/SObTemplate.bas#L26-L29
  [sob_tmp_f1]:   ../../../search?type=code&q=path:src/*Template.bas+content:FieldOne
  [sob_tmp_f2]:   ../../../search?type=code&q=path:src/*Template.bas+content:FieldTwo
  [sob_tmp_f3]:   ../../../search?type=code&q=path:src/*Template.bas+content:FieldThree
  [vbe_drop]:     https://stackoverflow.com/a/57894889
  [sob_tmp_fld]:  ../../../search?type=code&q=path:src/*Template.bas+content:*__Field.
  [sob_todo_9]:   ../src/SObTemplate.bas#L212
  [sob_tmp_acc]:  ../src/SObTemplate.bas#L171-L213
  [vba_prp]:      https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
  [vba_priv]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/private-statement
  [vba_prp_get]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [sob_tmp_get]:  ../src/SObTemplate.bas#L176-L178
  [vba_prp_let]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-let-statement
  [sob_tmp_scl]:  ../src/SObTemplate.bas#L180-L182
  [vba_prp_set]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-set-statement
  [sob_tmp_obj]:  ../src/SObTemplate.bas#L191-L193
  [sob_todo_3]:   ../src/SObTemplate.bas#L64
  [sob_tmp_ini]:  ../src/SObTemplate.bas#L65-L80
  [sob_flds]:     Fields.md
  [sob_todo_4]:   ../src/SObTemplate.bas#L99
  [sob_tmp_arr]:  ../src/SObTemplate.bas#L101-L106
  [sob_tmp_is]:   ../src/SObTemplate.bas#L89-L150
  [sob_todo_8]:   ../src/SObTemplate.bas#L162
  [sob_tmp_asn]:  ../src/SObTemplate.bas#L163-L166
  [sob_tmp_as]:   ../src/SObTemplate.bas#L154-L167
  [sob_todo_11]:  ../src/SObTemplate.bas#L286
  [sob_tmp_fmt]:  ../src/SObTemplate.bas#L277-L302
  [sob_todo_12]:  ../src/SObTemplate.bas#L290
  [sob_vis]:      Visualization.md
  [sob_todo_10]:  ../src/SObTemplate.bas#L245
  [sob_tmp_mtd]:  ../src/SObTemplate.bas#L217-L246
  [vba_fun]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/function-statement
  [vba_sub]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/sub-statement
  [sob_todo_5]:   ../src/SObTemplate.bas#L119
  [sob_tmp_p1]:   ../src/SObTemplate.bas#L175-L182
  [sob_tmp_chk]:  ../src/SObTemplate.bas#L111-L140
  [sob_todo_7]:   ../src/SObTemplate.bas#L148
  [sob_tmp_err]:  ../src/SObTemplate.bas#L149
  [sob_err_args]: Validation.md#syntax
  [vba_ppg_err]:  https://www.fastercapital.com/content/Error-Handling--Error-Handling-Excellence--Bulletproofing-Your-VBA-Code.html#Error-Bubbling-and-Propagation
  [sob_todo_6]:   ../src/SObTemplate.bas#L136
  [sob_tmp_vld]:  ../src/SObTemplate.bas#L130-L140
  [sob_tmp_cir]:  ../src/SObTemplate.bas#L108
