## Setup ##

Setup is quick and painless with [handy templates][sob_tmps].  Simply fill out the [`TODO`][sob_todos]s and paste the result in your module!


### Consolidated ###

To consolidate everything within your existing module, fill out [`SnippetTemplate.bas`][sob_snp_tmp] and paste into your module.  Then paste [`Snippet.bas`][sob_snp] alongside it.


### Dependency ###

To outsource the **`SOb`** framework to a single external dependency, fill out [`SObTemplate.bas`][sob_mod_tmp] and paste into your module.  Then instruct your end user to import the [`SOb.bas`][sob_mod] module, which you may reference as a [submodule][ghub_submod] in your repo.


> [!WARNING]
> 
> By outsourcing, you reduce your security against tampering!  The ["encryption"][sob_secure] is no longer scoped to your module, so others can overwrite[^9] your "private" fields.


### Template ###

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



  [sob_tmps]:     ../../search?type=code&q=path:src/*Template.bas
  [sob_todos]:    ../../search?type=code&q=path:src/*Template.bas+content:TODO:
  [sob_snp_tmp]:  src/SnippetTemplate.bas
  [sob_snp]:      src/Snippet.bas
  [sob_mod_tmp]:  src/SObTemplate.bas
  [sob_mod]:      src/SOb.bas
  [ghub_submod]:  https://github.blog/open-source/git/working-with-submodules
  [sob_secure]:   src/SOb.bas#L489-L504
  [sob_todo_1]:   src/SObTemplate.bas#L6
  [sob_tmp_ast]:  ../../search?type=code&q=path:src/*Template.bas+content:*
  [sob_cls]:      #typology
  [vba_const]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/declaring-constants
  [vba_enum]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [vba_proc]:     https://learn.microsoft.com/office/vba/language/how-to/create-a-procedure
  [vba_arg]:      https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
  [sob_todo_2]:   src/SObTemplate.bas#L25
  [sob_tmp_enm]:  src/SObTemplate.bas#L26-L29
  [sob_tmp_f1]:   ../../search?type=code&q=path:src/*Template.bas+content:FieldOne
  [sob_tmp_f2]:   ../../search?type=code&q=path:src/*Template.bas+content:FieldTwo
  [sob_tmp_f3]:   ../../search?type=code&q=path:src/*Template.bas+content:FieldThree
  [sob_fld]:      docs/Field.md
  [vbe_drop]:     https://stackoverflow.com/a/57894889
  [sob_tmp_fld]:  ../../search?type=code&q=path:src/*Template.bas+content:*__Field.
  [sob_todo_9]:   src/SObTemplate.bas#L212
  [sob_tmp_acc]:  src/SObTemplate.bas#L171-L213
  [vba_prp]:      https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
  [vba_priv]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/private-statement
  [vba_prp_get]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [sob_tmp_get]:  src/SObTemplate.bas#L176-L178
  [vba_prp_let]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-let-statement
  [sob_tmp_scl]:  src/SObTemplate.bas#L180-L182
  [vba_prp_set]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-set-statement
  [sob_tmp_obj]:  src/SObTemplate.bas#L191-L193
  [sob_todo_3]:   src/SObTemplate.bas#L64
  [sob_tmp_ini]:  src/SObTemplate.bas#L65-L80
  [sob_flds]:     docs/Fields.md
  [sob_todo_4]:   src/SObTemplate.bas#L99
  [sob_tmp_arr]:  src/SObTemplate.bas#L101-L106
  [sob_tmp_is]:   src/SObTemplate.bas#L89-L150
  [sob_todo_8]:   src/SObTemplate.bas#L162
  [sob_tmp_asn]:  src/SObTemplate.bas#L163-L166
  [sob_tmp_as]:   src/SObTemplate.bas#L154-L167
  [sob_todo_11]:  src/SObTemplate.bas#L286
  [sob_tmp_fmt]:  src/SObTemplate.bas#L277-L302
  [sob_todo_12]:  src/SObTemplate.bas#L290
  [sob_vis]:      docs/Visualization.md
  [sob_todo_10]:  src/SObTemplate.bas#L245
  [sob_tmp_mtd]:  src/SObTemplate.bas#L217-L246
  [vba_fun]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/function-statement
  [vba_sub]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/sub-statement
  [sob_todo_5]:   src/SObTemplate.bas#L119
  [sob_tmp_p1]:   src/SObTemplate.bas#L175-L182
  [sob_tmp_chk]:  src/SObTemplate.bas#L111-L140
  [sob_todo_7]:   src/SObTemplate.bas#L148
  [sob_tmp_err]:  src/SObTemplate.bas#L149
  [sob_err_args]: docs/Validation.md#syntax
  [vba_ppg_err]:  https://www.fastercapital.com/content/Error-Handling--Error-Handling-Excellence--Bulletproofing-Your-VBA-Code.html#Error-Bubbling-and-Propagation
  [sob_todo_6]:   src/SObTemplate.bas#L136
  [sob_tmp_vld]:  src/SObTemplate.bas#L130-L140
  [sob_tmp_cir]:  src/SObTemplate.bas#L108
  [sob_suite]:    #api
  [sob_print]:    #visualization
  [vba_cls]:      https://vbaplanet.com/objects.php
  [vba_udt]:      https://learn.microsoft.com/office/vba/language/how-to/user-defined-data-type
  [vba_cons]:     #old-problems
  [ghlp_repo]:    https://github.com/GregYannes/GitHelp#readme
  [so_post]:      https://codereview.stackexchange.com/q/293168
  [so_comm_1]:    https://codereview.stackexchange.com/posts/comments/583913
  [so_comm_2]:    https://codereview.stackexchange.com/posts/comments/584856
  [vba_cls_call]: https://stackoverflow.com/posts/comments/118407731
  [obj_cons]:     https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5995379
  [vb_bind]:      https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/early-late-binding
  [udt_cons]:     https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5971117
  [udt_silo]:     https://stackoverflow.com/q/38361276
  [udt_tamp]:     http://cpearson.com/excel/classes.aspx
  [udt_pass_var]: https://vbforums.com/showthread.php?304617-Storing-a-UDT-in-a-variant-type-mismatch#post1785101
  [udt_pass_obj]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5540423
  [udt_pass_clx]: https://vbforums.com/showthread.php?599355-RESOLVED-Addin-a-user-defined-type-to-a-collection
  [udt_pass_dix]: https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5971115
  [udt_dll]:      https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541509
  [udt_hack_srl]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5542053
  [udt_hack_prg]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541375
  [vba_clx]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_obj]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/object-data-type
  [vba_var]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [sob_setup]:    #setup
  [udt_lib]:      https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541458
  [vba_typ_lib]:  https://learn.microsoft.com/office/vba/language/how-to/set-reference-to-a-type-library
  [sob_typo]:     docs/Typology.md
  [vba_typ_fn]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/typename-function
  [vba_typ_op]:   https://learn.microsoft.com/dotnet/visual-basic/language-reference/operators/typeof-operator
  [vba_byref]:    https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [vb_net]:       https://learn.microsoft.com/dotnet/visual-basic
  [vba_tostring]: https://stackoverflow.com/posts/comments/98934630
  [net_tostring]: https://learn.microsoft.com/dotnet/fundamentals/runtime-libraries/system-object-tostring
  [sob_depend]:   #dependency
  [vba_opt_priv]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/option-private-statement
  [sob_meta]:     docs/Metadata.md
  [sem_ver]:      https://semver.org
  [sob_cre]:      docs/Creation.md
  [vba_arr_fn]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/array-function
  [sob_vali]:     docs/Validation.md
  [vba_prp_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-property-procedures
  [vbe_immed]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/immediate-window
  [vba_pub]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/public-statement
  [sob_util]:     docs/Utilities.md
  [vba_arr]:      https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_err_obj]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
