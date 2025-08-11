# SOb #

The **`SOb`** framework lets you simulate "objects" in VBA for Excel.

With a [full suite][sob_suite] of features at your fingertips — including [pretty printing][sob_print] — your <ins>**s**</ins>imulated <ins>**ob**</ins>ject ("**SOb**") can replicate an [object][vba_cls] or [UDT][vba_udt] without the [frustrating downsides][vba_cons].  No matter how many **SOb**s you need, or where you need them, this framework supports them within your _existing_ code.  No imports are needed!


## The **`SOb`** Story ##

I first encountered this use case when developing [**`GitHelp`**][ghlp_repo], which simulates a fielded "library" of documentation.  It demanded an [innovative approach][so_post], and as seasoned developers [chimed][so_comm_1] [in][so_comm_2], this took on a life of its own!

Like me, you might desire several such data structures, where _some_ fields are accessible (or not) to outside users.  These structures (like UDTs) are self-contained within your module, yet (like objects) they can be used by other modules.  Ideally these other modules should still compile in the absence of yours, which should be easy for lay users to (re)install.


### Old Problems ###

Unfortunately, neither objects nor UDTs achieve the outcome above!  For every object you include, your users must install an additional class module, and ["classes are a pain"][obj_cons] to develop.  Furthermore, their absence can derail compilation, unless other modules inefficiently resort to [late-binding][vb_bind].  But if objects are painful, then ["UDTs are _notoriously_ problematic"][udt_cons].

While UDTs are restrictively [siloed][udt_silo] within your module, their fields are still [vulnerable to editing][udt_tamp].  You cannot pass them to placeholders like [`Variant`][udt_pass_var] or [`Object`][udt_pass_obj], nor can you store them within a [`Collection`][udt_pass_clx] or (on Windows) a [`Dictionary`][udt_pass_dix].  To avoid burdening users with [prohibitive setup][udt_dll], developers have often resorted to [dubious][udt_hack_srl] [hacks][udt_hack_prg]!


### New Solution ###

The **`SOb`** framework addresses all these shortcomings.  It builds your SOb atop a [`Collection`][vba_clx], which is native to VBA across platforms (Windows and Mac).  You may let other modules access your SOb, yet its fields are ["encrypted"][sob_secure] against the more insidious tampering.  And while you may store your SObs as [`Object`][vba_obj]s, they require no class modules whatsoever—instead you can [easily set them all up][sob_setup] within your existing module!


## Setup ##

Setup is quick and painless with [handy templates][sob_tmpls].  Simply fill out the [`TODO`][sob_todos]s and paste the result in your module!


### Consolidated ###

To consolidate everything within your existing module, fill out [`SnippetTemplate.bas`][sob_snp_tmpl] and paste into your module.  Then paste [`Snippet.bas`][sob_snp] alongside it.


### Dependency ###

To outsource the **`SOb`** framework to a single external dependency, fill out [`SObTemplate.bas`][sob_mod_tmpl] and paste into your module.  Then instruct your end user to import the [`SOb.bas`][sob_mod] module, which you may reference as a [submodule][ghub_submod] in your repo.


### Template ###

Fill out either template according to these steps:

  1. [`TODO`][sob_todo_1]: Replace every [`*`][sob_tmpl_ast] with the (["class"][sob_cls]) name you desire for your SOb.  A simple **Find & Replace** should suffice.
     
     So if you name your SOb something like "`Foo`", this should yield a `FOO_CLASS` [constant][vba_const] atop your module; along with a `Foo__Field` [enumeration][vba_enum] below; followed by [procedures][vba_proc] of the form `Foo_…()` with a `foo` [argument][vba_arg].
     
  1. [`TODO`][sob_todo_2]: Enumerate the [fields][sob_tmpl_enm] you desire for your SOb.  Currently there are three placeholders for these fields: [`FieldOne`][sob_tmpl_f1] and [`FieldTwo`][sob_tmpl_f2] and [`FieldThree`][sob_tmpl_f3].  Feel free to **Find & Replace** these, and to append further fields of your own.
     
     This way, you can specify a field to (say) [`Obj_Field()`][sob_flds] by simply selecting it from the [dropdown][vbe_drop] for [`Foo__Field.…`][sob_tmpl_fld].
     
  1. [`TODO`][sob_todo_8]: Implement [accessors][sob_tmpl_acc] for your [`Foo__Field.…`][sob_tmpl_fld] fields.  Each should be a [`Property`][vba_prp] of the form `Foo_…(ByRef foo As Object)`, and you may restrict it to internal usage via the [`Private`][vba_priv] statement.
     
     Simply wrap [`Obj_Get()`][sob_flds] with a [`Property Get`][vba_prp_get] to [_retrieve_][sob_tmpl_get] a field...
     
     ```vba
     Property Get Foo_FieldOne(ByRef foo As Object) As Integer
     	Obj_Get Foo_FieldOne, foo, Foo__Field.FieldOne
     End Property
     ```
     
     ...but implement a [`Property Let`][vba_prp_let] to _update_ a [scalar field][sob_tmpl_scl]...
     
     ```vba
     Property Let Foo_FieldOne(ByRef foo As Object, ByVal val As Integer)
     	Let Foo_FieldOne = val
     End Property
     ```
     
     ...or a [`Property Set`][vba_prp_set] to update an [objective field][sob_tmpl_obj].
     
     ```vba
     Property Set Foo_FieldTwo(ByRef foo As Object, ByRef val As Range)
     	Set Foo_FieldTwo = val
     End Property
     ```
     
  1. [`TODO`][sob_todo_3]: Initialize the values for your fields, within [`Foo_Initialize()`][sob_tmpl_ini].  Use [`Obj_HasField()`][sob_flds] to test whether a field exists, and when it does not, use your [accessor][sob_tmpl_acc] to set its initial value.
     
  1. [`TODO`][sob_todo_4]: List all your [`Foo__Field.…`][sob_tmpl_fld] fields in the [`Array(...)`][sob_tmpl_arr] call, within [`IsFoo()`][sob_tmpl_is].
     
     This way, `IsFoo()` checks that a "Foo" object has all its fields.
     
  1. [`TODO`][sob_todo_7]: Using your accessors, [assign each field][sob_tmpl_asn] from `obj` to its corresponding field in `AsFoo`, within [`AsFoo()`][sob_tmpl_as].
     
     This way, `AsFoo()` can coerce any input (`x`) into a "Foo" object, by extracting fields from the former into the latter.
     
  1. [`TODO`][sob_todo_10]: Create any summary (`sum`) or detail (`dtl`) you wish, to visually represent your object within [`Foo_Format()`][sob_tmpl_fmt].  The **`SOb`** framework _automatically_ formats these for you: summaries display on a single line...
     
     > ```
     > <Foo[sum]>
     > ```
     
     ...and details display across multiple lines:
     
     > ```
     > <Foo: {
     > 	dtl_1
     > 	dtl_2
     > 	…
     > }>
     > ```
     
  1. [`TODO`][sob_todo_11]: Pass any summary (`sum`) or detail (`dtl`) to [`Obj_Format()`][sob_vis], along with all arguments from [`Foo_Format()`][sob_tmpl_fmt].
     
     This way, others can apply various settings when printing your "Foo" object, including developers who wish to build their own SObs upon "Foo".

You may _optionally_ enhance "Foo" with further steps:

  9. [`TODO`][sob_todo_9]: Implement any [methods][sob_tmpl_mtd] you desire, which operate on your "Foo" object.  Each should be a [`Function`][vba_fun] or [`Sub`routine][vba_sub] of the form `Foo_…(ByRef foo As Object, …)` where `foo` is followed by any [arguments][vba_arg] needed by the method.  You may restrict it to internal usage via the [`Private`][vba_priv] statement.
     
  9. [`TODO`][sob_todo_5]: Call all your field accessors like [`Foo_FieldOne()`][sob_tmpl_p1], in the [`Check ...`][sob_tmpl_chk] call within [`IsFoo()`][sob_tmpl_is].
     
     This way, `IsFoo()` also checks that the "Foo" fields are of the expected type, and so forth.
     
  9. [`TODO`][sob_todo_6]: Customize any [further validation][sob_tmpl_vld] you wish [`IsFoo()`][sob_tmpl_is] to perform.  Each validation step should assign a `Boolean` value to `IsFoo`...
     
     ```vba
     	IsFoo = ...
     ```
     
     ...and finish by [short-circuiting][sob_tmpl_cir] when `False`.
     
     ```vba
     	If Not IsFoo Then Exit Function
     ```

Now you are ready to work with "Foo" objects, within your module and elsewhere!


## API ##

Here are all the features provided by **`SOb`** for developers.  To avoid confusing _your_ users, the [**`SOb`** module][sob_mod] hides its own functions from Excel, via [`Option Private`][vba_opt_priv].


### Metadata ###

Describe the [**`SOb`** module][sob_mod] _itself_.

  - [`MOD_NAME`][sob_meta]: The name (`String`) of the module: `"SOb"`.
  - [`MOD_VERSION`][sob_meta]: Its current [version][sem_ver] (`String`): `"0.1.0"`.
  - [`MOD_REPO`][sob_meta]: The URL (`String`) to its repo: `"https://github.com/GregYannes/SOb"`


### Creation ###

"Declare" a new SOb.

  - [`New_Obj()`][sob_cre]: Returns an initialized SOb (`Object`).
  - [`Obj_Initialize()`][sob_cre]: Initializes a generic `Object` as an SOb.


### Typology ###

Ascertain the "type" of an SOb...

  - [`Obj_Class()`][sob_typo]: Retrieve the simulated "class" (`String`) of an SOb.
  - [`IsObj()`][sob_typo]: Test (`Boolean`) if something is an SOb.

...and manipulate that type.

  - [`AsObj()`][sob_typo]: Cast something as an SOb (`Object`).


### Fields ###

Access simulated "fields" in an SOb...

  - [`Obj_Field()`][sob_flds]: Read ([`Get`][vba_prp_get]) and write ([`Let`][vba_prp_let] or [`Set`][vba_prp_set]) the field as a [`Property`][vba_prp].
  - [`Obj_Get()`][sob_flds]: A delegate of [`Property Get`][vba_prp_get] with protection against missing fields.

...along with metadata about such fields.

  - [`Obj_FieldCount()`][sob_flds]: The (maximum) count (`Long`) of simulated fields in an SOb.
  - [`Obj_HasField()`][sob_flds]: Test (`Boolean`) if an SOb has a certain field.
  - [`Obj_HasFields()`][sob_flds]: Test (`Boolean`) if an SOb has an entire set of fields, wrapped in an [`Array()`][vba_arr_fn]...
  - [`Obj_HasFields0()`][sob_flds]: ...or entered manually.


### Validation ###

Validate SObs within advanced implementations of [`Is*()`][sob_tmpl_chk].

  - [`Obj_Check()`][sob_vali]: [Call][vba_prp_call] your [accessors][sob_tmpl_acc] without assignment, merely to test (say) their type integrity.
  - [`Obj_CheckError()`][sob_vali]: Test (`Boolean`) if _certain_ errors (like type) invalidate the check, but propagate any _other_ errors.


### Visualization ###

Textually visualize the entire SOb...

  - [`Obj_Print()`][sob_vis]: Print (`String`) an SOb to the [console][vbe_immed] with automatic formatting.
  - [`Obj_Print0()`][sob_vis]: Print something (`String`) verbatim to the console.
  - [`Obj_Format()`][sob_vis]: Automatically format (`String`) an SOb for printing.

...or specifically its fields in detail.

  - [`Obj_FormatFields()`][sob_vis]: Automatically format (`String`) a set of simulated fields, wrapped in an [`Array()`][vba_arr_fn]...
  - [`Obj_FormatFields0()`][sob_vis]: ...or entered manually with default settings.


### Utilities ###

Perform broadly useful ([`Public`][vba_pub]) tasks via the [**`SOb`** module][sob_mod]...

  - [`Assign()`][sob_util]: Assign any value (scalar or objective) to a variable (by [reference][vba_byref]).
  - [`Txt_Indent()`][sob_util]: Indent (`String`) some lines of text.

…along with further ([`Private`][vba_priv]) tasks via an [**`SOb`** snippet][sob_snp] in your own module.

  - [`Clx_Has()`][sob_util]: Test (`Boolean`) if a [`Collection`][vba_clx] contains an item.
  - [`Clx_Get()`][sob_util]: Safely retrieve any item (`Variant`) from a `Collection`.
  - [`Clx_Set()`][sob_util]: Set the value of an item in a `Collection`.
  - [`Arr_Length()`][sob_util]: Get the length (`Long`) of an [array][vba_arr].
  - [`Err_Raise()`][sob_util]: Raise an [error object][vba_err_obj] directly.
  - [`Txt_Contains()`][sob_util]: Test (`Boolean`) if text contains a substring.



  [sob_suite]:    #api
  [sob_print]:    #visualization
  [vba_cls]:      https://vbaplanet.com/objects.php
  [vba_udt]:      https://learn.microsoft.com/office/vba/language/how-to/user-defined-data-type
  [vba_cons]:     #old-problems
  [ghlp_repo]:    https://github.com/GregYannes/GitHelp#readme
  [so_post]:      https://codereview.stackexchange.com/q/293168
  [so_comm_1]:    https://codereview.stackexchange.com/posts/comments/583913
  [so_comm_2]:    https://codereview.stackexchange.com/posts/comments/584856
  [obj_cons]:     https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5995379
  [vb_bind]:      https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/early-late-binding
  [udt_cons]:     https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5971117
  [udt_silo]:     https://stackoverflow.com/a/41689531
  [udt_tamp]:     http://cpearson.com/excel/classes.aspx
  [udt_pass_var]: https://vbforums.com/showthread.php?304617-Storing-a-UDT-in-a-variant-type-mismatch#post1785101
  [udt_pass_obj]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5540423
  [udt_pass_clx]: https://vbforums.com/showthread.php?599355-RESOLVED-Addin-a-user-defined-type-to-a-collection
  [udt_pass_dix]: https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5971115
  [udt_dll]:      https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541509
  [udt_hack_srl]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5542053
  [udt_hack_prg]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541375
  [vba_clx]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [sob_secure]:   src/SOb.bas#L489-L504
  [vba_obj]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/object-data-type
  [sob_setup]:    #setup
  [sob_tmpls]:    ../../search?type=code&q=path:src/*Template.bas
  [sob_todos]:    ../../search?type=code&q=path:src/*Template.bas+content:TODO:
  [sob_snp_tmpl]: src/SnippetTemplate.bas
  [sob_snp]:      src/Snippet.bas
  [sob_mod_tmpl]: src/SObTemplate.bas
  [sob_mod]:      src/SOb.bas
  [ghub_submod]:  https://github.blog/open-source/git/working-with-submodules
  [sob_todo_1]:   src/SObTemplate.bas#L6
  [sob_tmpl_ast]: ../../search?type=code&q=path:src/*Template.bas+content:*
  [sob_cls]:      #typology
  [vba_const]:    https://learn.microsoft.com/office/vba/language/concepts/getting-started/declaring-constants
  [vba_enum]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [vba_proc]:     https://learn.microsoft.com/office/vba/language/how-to/create-a-procedure
  [vba_arg]:      https://learn.microsoft.com/office/vba/language/concepts/getting-started/understanding-named-arguments-and-optional-arguments
  [sob_todo_2]:   src/SObTemplate.bas#L25
  [sob_tmpl_enm]: src/SObTemplate.bas#L26-L29
  [sob_tmpl_f1]:  ../../search?type=code&q=path:src/*Template.bas+content:FieldOne
  [sob_tmpl_f2]:  ../../search?type=code&q=path:src/*Template.bas+content:FieldTwo
  [sob_tmpl_f3]:  ../../search?type=code&q=path:src/*Template.bas+content:FieldThree
  [sob_flds]:     docs/Fields.md
  [vbe_drop]:     https://stackoverflow.com/a/57894889
  [sob_tmpl_fld]: ../../search?type=code&q=path:src/*Template.bas+content:*__Field.
  [sob_todo_8]:   src/SObTemplate.bas#L211
  [sob_tmpl_acc]: src/SObTemplate.bas#L170-L212
  [vba_prp]:      https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
  [vba_priv]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/private-statement
  [vba_prp_get]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [sob_tmpl_get]: src/SObTemplate.bas#L175-L177
  [vba_prp_let]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-let-statement
  [sob_tmpl_scl]: src/SObTemplate.bas#L179-L181
  [vba_prp_set]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-set-statement
  [sob_tmpl_obj]: src/SObTemplate.bas#L190-L192
  [sob_todo_3]:   src/SObTemplate.bas#L64
  [sob_tmpl_ini]: src/SObTemplate.bas#L65-L80
  [sob_todo_4]:   src/SObTemplate.bas#L99
  [sob_tmpl_arr]: src/SObTemplate.bas#L101-L106
  [sob_tmpl_is]:  src/SObTemplate.bas#L89-L149
  [sob_todo_7]:   src/SObTemplate.bas#L161
  [sob_tmpl_asn]: src/SObTemplate.bas#L162-L165
  [sob_tmpl_as]:  src/SObTemplate.bas#L153-L166
  [sob_todo_10]:  src/SObTemplate.bas#L285
  [sob_tmpl_fmt]: src/SObTemplate.bas#L276-L301
  [sob_todo_11]:  src/SObTemplate.bas#L289
  [sob_vis]:      docs/Visualization.md
  [sob_todo_9]:   src/SObTemplate.bas#L244
  [sob_tmpl_mtd]: src/SObTemplate.bas#L216-L245
  [vba_fun]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/function-statement
  [vba_sub]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/sub-statement
  [sob_todo_5]:   src/SObTemplate.bas#L119
  [sob_tmpl_p1]:  src/SObTemplate.bas#L174-L181
  [sob_tmpl_chk]: src/SObTemplate.bas#L111-L140
  [sob_todo_6]:   src/SObTemplate.bas#L136
  [sob_tmpl_vld]: src/SObTemplate.bas#L130-L140
  [sob_tmpl_cir]: src/SObTemplate.bas#L108
  [vba_opt_priv]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/option-private-statement
  [sob_meta]:     docs/Metadata.md
  [sem_ver]:      https://semver.org
  [sob_cre]:      docs/Creation.md
  [sob_typo]:     docs/Typology.md
  [vba_arr_fn]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/array-function
  [sob_vali]:     docs/Validation.md
  [vba_prp_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-property-procedures
  [vbe_immed]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/immediate-window
  [vba_pub]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/public-statement
  [sob_util]:     docs/Utilities.md
  [vba_byref]:    https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [vba_arr]:      https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_err_obj]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
