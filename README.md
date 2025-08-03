# SOb #

The **`SOb`** framework lets you simulate "objects" in Excel VBA.

With a [full suite][sob_suite] of features at your fingertips — including [pretty printing][sob_print] — your simulated object ("**SOb**") can replicate an [object][vba_cls] or [UDT][vba_udt] without the [frustrating downsides][vba_cons].  No matter how many **SOb**s you need, or where you need them, this framework supports them within your _existing_ code.  No imports are needed!


## The **`SOb`** Story ##

I first encountered this use case when developing [**`GitHelp`**][gh_repo], which simulates a fielded "library" of documentation.

Like me, you might desire several such data structures, where _some_ fields are accessible (or not) to outside users.  These structures (like UDTs) are self-contained within your module, yet (like objects) they can be used by other modules.  Ideally these other modules should still compile in the absence of yours, which should be easy for lay users to (re)install.


### Old Problems ###

Unfortunately, neither objects nor UDTs achieve the outcome above!  Every object requires users to install an extra class module, and ["classes are a pain"][obj_cons] to develop.  Furthermore, their absence can derail compilation, unless other modules inefficiently resort to [late-binding][vb_bind].  But if objects are painful, then ["UDTs are _notoriously_ problematic"][udt_cons].

While UDTs are generally [siloed][udt_silo] within your module, their fields are still [vulnerable to editing][udt_tamp].  You cannot pass them to placeholders like [`Variant`][udt_pass_var] or [`Object`][udt_pass_obj], nor can you store them within a [`Collection`][udt_pass_clx] or (on Windows) a [`Dictionary`][udt_pass_dix].  To avoid burdening users with [prohibitive setup][udt_dll], developers have often resorted to [dubious][udt_hack_srl] [hacks][udt_hack_prg]!


### New Solution ###

The **`SOb`** framework addresses all these shortcomings.  It builds your SOb atop a [`Collection`][vba_clx], which is native to VBA across platforms (Windows and Mac).  You may let other modules access your SOb, yet its fields are ["encrypted"][sob_secure] against the more insidious tampering.  And while you may store your SObs as [`Object`][vba_obj]s, they require no class modules whatsoever—instead you can [easily set them up][sob_setup] all within your existing module!


## Setup ##

![](med/banner_unfinished.png)


## API ##

Here are all the features provided by **`SOb`** for developers.  To avoid confusing _your_ users, any **`SOb`** module hides its own functions from Excel, via [`Option Private`][vba_opt_priv].


### Metadata ###

Describe the **`SOb`** module _itself_.

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

Access simulated fields in an SOb...

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
  - [`Obj_Error()`][sob_vali]: Test (`Boolean`) if certain errors (like type) invalidate the check, but propagates other errors.


### Visualization ###

Textually visualize the entire object...

  - [`Obj_Print()`][sob_vis]: Print (`String`) an SOb to the [console][vba_immed] with automatic formatting.
  - [`Obj_Print0()`][sob_vis]: Print something (`String`) verbatim to the console.
  - [`Obj_Format()`][sob_vis]: Automatically format (`String`) an SOb for printing.

...or specifically its fields in detail.

  - [`Obj_FormatFields()`][sob_vis]: Automatically format (`String`) a set of simulated fields, wrapped in an [`Array()`][vba_arr_fn]...
  - [`Obj_FormatFields0()`][sob_vis]: ...or entered manually with default settings.


### Utilities ###

Perform broadly useful ([`Public`][vba_pub]) tasks via an **`SOb`** module...

  - [`Assign()`][sob_util]: Assign any value (scalar or objective) to a variable (by [reference][vba_byref]).
  - [`Txt_Indent()`][sob_util]: Indent (`String`) some lines of text.

...along with further ([`Private`][vba_priv]) tasks via an SOb snippet in your own module.

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
  [gh_repo]:      https://github.com/GregYannes/GitHelp
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
  [sob_secure]:   src/SOb.bas#L479-L494
  [vba_obj]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/object-data-type
  [sob_setup]:    #setup
  [vba_opt_priv]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/option-private-statement
  [sob_meta]:     docs/Metadata.md
  [sem_ver]:      https://semver.org
  [sob_cre]:      docs/Creation.md
  [sob_typo]:     docs/Typology.md
  [sob_flds]:     docs/Fields.md
  [vba_prp_get]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [vba_prp_let]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-let-statement
  [vba_prp_set]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-set-statement
  [vba_prp]:      https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
  [vba_arr_fn]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/array-function
  [sob_tmpl_chk]: src/Template.bas#L111-L140
  [sob_vali]:     docs/Validation.md
  [vba_prp_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-property-procedures
  [sob_tmpl_acc]: src/Template.bas#L170-L212
  [sob_vis]:      docs/Visualization.md
  [vba_immed]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/immediate-window
  [vba_pub]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/public-statement
  [sob_util]:     docs/Utilities.md
  [vba_byref]:    https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [vba_priv]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/private-statement
  [vba_arr]:      https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_err_obj]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
