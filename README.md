# SOb #

The **`SOb`** framework lets you simulate "objects" in VBA.

With a [full suite][sob_suite] of features at your fingertips — including [pretty printing][sob_print] — your <ins>**s**</ins>imulated <ins>**ob**</ins>ject ("**SOb**") mimics an [object][vba_cls] or [UDT][vba_udt] without the [frustrating downsides][vba_cons].  No matter how many **SOb**s you need, or where you need them, this framework supports them within your _existing_ code.  No imports are needed!


## The **`SOb`** Story ##

I first encountered this use case when developing [**`GitHelp`**][ghlp_repo], which simulates a fielded "library" of documentation.  It demanded an [innovative approach][so_post], and as seasoned developers [chimed][so_comm_1] [in][so_comm_2], this took on a life of its own!

Like me, you might desire several such data structures, where _some_ fields are accessible (or not) to outside users.  These structures (like UDTs) are self-contained within your module, yet (like objects) they can be used by [object classes][vba_cls] and modules alike.  Ideally these other modules should still compile in the absence of yours, which should be easy for lay users to (re)install.

Unfortunately, neither objects nor UDTs achieve this outcome!  For every object you include, your users must install an additional class module.  And if objects ["are a pain"][obj_cons], then ["UDTs are _notoriously_ problematic"][udt_cons].


## Advantages ##

The **`SOb`** framework addresses all these shortcomings.  It builds your SOb atop a [`Collection`][vba_clx], which is native to VBA across platforms (Windows and Mac).  And unlike classes, your SObs carry no baggage whatsoever—you can [easily set them _all_ up][sob_setup] within your existing module!

| <ins>Feature</ins> | <ins>Description</ins>                                                            | <ins>SOb</ins> | <ins>Object</ins> | <ins>UDT</ins> |
| :----------------- | :-------------------------------------------------------------------------------- | :------------- | :---------------- | :------------- |
| Painless           | Is it quick and easy for you to code?                                             | ✓              |   [^1]            | ✓              |
| Installable        | Is it quick and easy for lay _users_ to install your code?                        | ✓              |   [^2]            | ✓              |
| Native             | Is it native to VBA?                                                              | ✓              | ✓                 | ✓              |
| Portable           | Does it work across all platforms?                                                | ✓              | ✓                 | ✓              |
| Independent        | Is it free of external dependencies?                                              | ✓              |   [^3]            | ✓              |
| Global             | Can it be used seamlessly across other modules and classes?                       | ✓ [^4]         | ✓                 |   [^5]         |
| Compilation        | Can its dependents compile in its absence?                                        | ✓              |   [^6]            |   [^7]         |
| Instantiation      | Can you dynamically declare new instances _after_ design time?                    | ✓              | ✓ [^a]            |   [^d]         |
| Placeholder        | Can it be passed to a generic [`Variant`][vba_var] or [`Object`][vba_obj]?        | ✓              | ✓ [^9]            |   [^7]         |
| Collectible        | Can it be included within a [`Collection`][vba_clx] (or [`Dictionary`][vba_dix])? | ✓              | ✓ [^10]           |   [^7]         |
| Identity           | Is its type identifiable by name, so you can distinguish it?                      | ✓ [^11]        | ✓ [^12]           |                |
| Methods            | Does it support [procedures][vba_proc] that operate on it?                        | ✓ [^13]        | ✓ [^f]            |   [^b][^14]    |
| Printing           | Does it support pretty printing for visualization?                                | ✓              |   [^15]           |                |
| Validation         | Can it validate values before they are assigned to fields?                        | ✓ [^16]        | ✓ [^c]            |   [^e]         |
| Private            | Can you hide certain fields (and "methods") from your user?                       | ✓ [^17]        | ✓ [^17]           |                |
| Secure             | Are its fields secure against unauthorized editing?                               | ✓ [^18]        | ✓                 |   [^8]         |


## Setup ##

Setup is quick and painless with [handy templates][sob_tmps].  Simply fill out the [`TODO`][sob_todos]s and paste the result in your module!  See [here][sob_doc_sup] for detailed instructions.

  - To consolidate everything within your module, use [`SnippetTemplate.bas`][sob_snp_tmp] alongside [`Snippet.bas`][sob_snp].  See [here][sob_consld] for details.
  - To outsource the **`SOb`** framework to a dependency, use [`SObTemplate.bas`][sob_mod_tmp] but import [`SOb.bas`][sob_mod] separately.  See [here][sob_outsrc] for details.


## Usage ##

Using an SOb is analogous to using an object.  The **`SOb`** framework provides a [backend][sob_suite], which lets you [implement][sob_sup_tmp] your frontend for your actual SOb.

Simply [enumerate][vba_enum] its fields (like "`Bar`") in the [template][sob_tmp_enm], and you may manipulate your SOb ("`Foo`") as illustrated below.

```vba
Private Enum Foo__Fields
	Bar
	' ...
End Enum
```

<br>

See [documentation][sob_docs] for further details and concrete [examples][sob_doc_ex].

| <ins>Action</ins> |   | <ins>Frontend</ins> | <ins>Backend</ins>       |   | <ins>Object</ins> |   | <ins>UDT</ins> |
| :---------------- | - | :------------------ | :----------------------- | - | :---------------- | - | :------------- |
| Declaration       |   | `Dim x As Object`   | `Dim x As Object`        |   | `Dim x As Foo`    |   | `Dim x As Foo` |
| Instantiation     |   | `Set x = New_Foo()` | `Set x = New_Obj("Foo")` |   | `Set x = New Foo` |   |                |
| Reading           |   | `Foo_Bar(x)`        | `Obj_Field(x, Bar)`      |   | `x.Bar`           |   | `x.Bar`        |
| Writing           |   | `Foo_Bar(x) = 1`    | `Obj_Field(x, Bar) = 1`  |   | `x.Bar = 1`       |   | `x.Bar = 1`    |
| Invocation        |   | `Foo_Fun(x, …)`     |                          |   | `x.Fun(…)`        |   |                |


## API ##

Here are all the features provided by **`SOb`** for developers.  To avoid confusing _your_ users, the [**`SOb`** module][sob_mod] hides its own functions from Excel, via [`Option Private`][vba_opt_priv].


### Metadata ###

Describe the [**`SOb`** module][sob_mod] _itself_.

  - [`MOD_NAME`][sob_meta]: The name (`String`) of the module.
  - [`MOD_VERSION`][sob_meta]: Its current [version][sem_ver] (`String`).
  - [`MOD_REPO`][sob_meta]: The URL (`String`) to its repository.


### Creation ###

"Declare" a new SOb.

  - [`New_Obj()`][sob_cre]: Returns an initialized SOb (`Object`).
  - [`Obj_Initialize()`][sob_cre]: Initializes a generic `Object` as an SOb.


### Typology ###

Ascertain the "type" of an SOb…

  - [`Obj_Class()`][sob_typo]: Retrieve the simulated "class" (`String`) of an SOb.
  - [`IsObj()`][sob_typo]: Test (`Boolean`) if something is an SOb.

…and manipulate that type.

  - [`AsObj()`][sob_typo]: Cast something as an SOb (`Object`).


### Fields ###

Access simulated "fields" in an SOb…

  - [`Obj_Field()`][sob_fld]: Read ([`Get`][vba_prp_get]) and write ([`Let`][vba_prp_let] or [`Set`][vba_prp_set]) the field as a [`Property`][vba_prp].
  - [`Obj_Get()`][sob_fld]: A delegate of [`Property Get`][vba_prp_get] with protection against missing fields.

…along with metadata about such fields.

  - [`Obj_FieldCount()`][sob_flds]: The (maximum) count (`Long`) of simulated fields in an SOb.
  - [`Obj_HasField()`][sob_flds]: Test (`Boolean`) if an SOb has a certain field.
  - [`Obj_HasFields()`][sob_flds]: Test (`Boolean`) if an SOb has an entire set of fields, wrapped in an [`Array()`][vba_arr_fn]…
  - [`Obj_HasFields0()`][sob_flds]: …or entered manually.


### Validation ###

Validate SObs within advanced implementations of [`Is*()`][sob_tmp_chk].

  - [`Obj_Check()`][sob_vali]: [Call][vba_prp_call] your [accessors][sob_tmp_acc] without assignment, merely to test (say) their type integrity.
  - [`Obj_CheckError()`][sob_vali]: Test (`Boolean`) if _certain_ errors (like type) invalidate the check, but propagate any _other_ errors.


### Visualization ###

Textually visualize the entire SOb…

  - [`Obj_Print()`][sob_vis]: Print (`String`) an SOb to the [console][vbe_immed] with automatic formatting.
  - [`Obj_Print0()`][sob_vis]: Print something (`String`) verbatim to the console.
  - [`Obj_Format()`][sob_vis]: Automatically format (`String`) an SOb for printing.

…or specifically its fields in detail.

  - [`Obj_FormatFields()`][sob_vis]: Automatically format (`String`) a set of simulated fields, wrapped in an [`Array()`][vba_arr_fn]…
  - [`Obj_FormatFields0()`][sob_vis]: …or entered manually with default settings.


### Utilities ###

Perform broadly useful ([`Public`][vba_pub]) tasks via the [**`SOb`** module][sob_mod]…

  - [`Assign()`][sob_util]: Assign any value (scalar or objective) to a variable (by [reference][vba_byref]).
  - [`Txt_Indent()`][sob_util]: Indent (`String`) some lines of text.

…along with further ([`Private`][vba_priv]) tasks via an [**`SOb`** snippet][sob_snp] in your own module.

  - [`Clx_Has()`][sob_util]: Test (`Boolean`) if a [`Collection`][vba_clx] contains an item.
  - [`Clx_Get()`][sob_util]: Safely retrieve any item (`Variant`) from a `Collection`.
  - [`Clx_Set()`][sob_util]: Set the value of an item in a `Collection`.
  - [`Arr_Length()`][sob_util]: Get the length (`Long`) of an [array][vba_arr].
  - [`Err_Raise()`][sob_util]: Raise an [error object][vba_err_obj] directly.
  - [`Txt_Contains()`][sob_util]: Test (`Boolean`) if text contains a substring.



  [^1]:  ["Classes are a pain"][obj_cons] to develop.
  [^2]:  To avoid burdening users with [prohibitive setup][udt_dll], developers have often resorted to [dubious][udt_hack_srl] [hacks][udt_hack_prg]!
  [^3]:  For every object you include, your users must install an additional class module.
  [^4]:  A class module may [call procedures from standard modules][vba_cls_call], like your own module or even the [**`SOb`** module][sob_mod].
  [^5]:  UDTs are restrictively [siloed][udt_silo] between classes and modules.
    
    There is only one exception[^7].
  [^6]:  Their absence can derail compilation, unless other modules inefficiently resort to [late-binding][vb_bind].
  [^7]:  Not unless you [reference the UDT][udt_lib] in a [type library][vba_typ_lib].
  [^8]:  Their fields are still [vulnerable to editing][udt_tamp].
  [^9]:  You cannot pass them to placeholders like [`Variant`][udt_pass_var] or [`Object`][udt_pass_obj]…
  [^10]: …nor can you include them within a [`Collection`][udt_pass_clx] or (on Windows) a [`Dictionary`][udt_pass_dix].
  [^11]: Via [`Obj_Class()`][sob_typo] and [`IsObj()`][sob_typo].
  [^12]: Via the [`TypeName()`][vba_typ_fn] function or the [`TypeOf`][vba_typ_op] operator.
  [^13]: Technically these ["methods"][sob_tmp_mtd] are simply modular [procedures][vba_proc] of the form `SOb_Method(sob, …)`, where the `sob` is passed [by reference][vba_byref].
  [^14]: Technically you _could_ mimic an SOb and implement "methods"[^13] of the form `UDT_Method(udt, …)`, where the `udt` is passed [by reference][vba_byref].
  [^15]: Unlike [.NET][vb_net] and other languages, VBA [does not implement][vba_tostring] a prototypical [`.ToString()`][net_tostring] method for objects.
  [^16]: The [accessors][sob_tmp_acc] for your SOb are [`Property`][vba_prp_set] procedures, in which you may validate input before assigning it to the field.
  [^17]: Via the [`Private`][vba_priv] keyword for [properties][vba_prp] (and [procedures][vba_proc]).
  [^18]: Its fields are ["encrypted"][sob_secure] against the more insidious tampering.  Others cannot typically _overwrite_ the value of a "private" field in your SOb—though they can _remove_ the field, which effectively resets it to an uninitialized state.
    
    However, if you [outsource the framework][sob_outsrc] from your module to the [**`SOb`** module][sob_mod], then others _can_ overwrite it via [`SOb.Obj_Field()`][sob_fld].
  [^a]:  You [may declare][obj_inst] new instances of objects at runtime, using the [`New`][vba_new] keyword…
  [^d]:  …but you [may not declare][udt_inst] new instances of UDTs.
  [^f]:  Objects [support][obj_act] methods for performing actions…
  [^b]:  …but UDTs [do not support][udt_inact] methods and ["cannot carry out actions"][udt_inact].
  [^c]:  Objects use [`Property`][vba_prp_set] procedures to [validate][obj_valid] values for fields…
  [^e]:  …but UDTs have [no mechanism][udt_tamp] for validation.



  [sob_suite]:    #api
  [sob_print]:    #visualization
  [vba_cls]:      https://vbaplanet.com/objects.php
  [vba_udt]:      https://learn.microsoft.com/office/vba/language/how-to/user-defined-data-type
  [vba_cons]:     #advantages
  [ghlp_repo]:    https://github.com/GregYannes/GitHelp#readme
  [so_post]:      https://codereview.stackexchange.com/q/293168
  [so_comm_1]:    https://codereview.stackexchange.com/posts/comments/583913
  [so_comm_2]:    https://codereview.stackexchange.com/posts/comments/584856
  [obj_cons]:     https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5995379
  [udt_cons]:     https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5971117
  [vba_clx]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [sob_setup]:    #setup
  [vba_var]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/variant-data-type
  [vba_obj]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/object-data-type
  [vba_dix]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/dictionary-object
  [vba_proc]:     https://learn.microsoft.com/office/vba/language/how-to/create-a-procedure
  [sob_tmps]:     ../../search?type=code&q=path:src/*Template.bas
  [sob_todos]:    ../../search?type=code&q=path:src/*Template.bas+content:TODO:
  [sob_doc_sup]:  docs/Setup.md
  [sob_snp_tmp]:  src/SnippetTemplate.bas
  [sob_snp]:      src/Snippet.bas
  [sob_consld]:   docs/Setup.md#consolidated
  [sob_mod_tmp]:  src/SObTemplate.bas
  [sob_mod]:      src/SOb.bas
  [sob_outsrc]:   docs/Setup.md#outsourced
  [sob_sup_tmp]:  docs/Setup.md#template
  [vba_enum]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [sob_tmp_enm]:  ../src/SObTemplate.bas#L26-L29
  [sob_docs]:     docs/
  [sob_doc_ex]:   ../../search?type=code&q=path:docs/*.md+content:%2F^%23%2B%5Cs%2BExamples%5Cs%2B%23%2B$%2F
  [vba_opt_priv]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/option-private-statement
  [sob_meta]:     docs/Metadata.md
  [sem_ver]:      https://semver.org
  [sob_cre]:      docs/Creation.md
  [sob_typo]:     docs/Typology.md
  [sob_fld]:      docs/Field.md
  [vba_prp_get]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-get-statement
  [vba_prp_let]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-let-statement
  [vba_prp_set]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/property-set-statement
  [vba_prp]:      https://learn.microsoft.com/office/vba/language/glossary/vbe-glossary#property
  [sob_flds]:     docs/Fields.md
  [vba_arr_fn]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/array-function
  [sob_tmp_chk]:  src/SObTemplate.bas#L111-L140
  [sob_vali]:     docs/Validation.md
  [vba_prp_call]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/calling-property-procedures
  [sob_tmp_acc]:  src/SObTemplate.bas#L171-L213
  [sob_vis]:      docs/Visualization.md
  [vbe_immed]:    https://learn.microsoft.com/office/vba/language/reference/user-interface-help/immediate-window
  [vba_pub]:      https://learn.microsoft.com/office/vba/language/reference/user-interface-help/public-statement
  [sob_util]:     docs/Utilities.md
  [vba_byref]:    https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [vba_priv]:     https://learn.microsoft.com/office/vba/language/reference/user-interface-help/private-statement
  [vba_arr]:      https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_err_obj]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/err-object
  [udt_dll]:      https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541509
  [udt_hack_srl]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5542053
  [udt_hack_prg]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541375
  [vba_cls_call]: https://stackoverflow.com/posts/comments/118407731
  [udt_silo]:     https://stackoverflow.com/q/38361276
  [vb_bind]:      https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/early-late-binding
  [udt_lib]:      https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5541458
  [vba_typ_lib]:  https://learn.microsoft.com/office/vba/language/how-to/set-reference-to-a-type-library
  [udt_tamp]:     http://cpearson.com/excel/classes.aspx
  [udt_pass_var]: https://vbforums.com/showthread.php?304617-Storing-a-UDT-in-a-variant-type-mismatch#post1785101
  [udt_pass_obj]: https://vbforums.com/showthread.php?893813-Passing-UDT-as-variant-for-saving-loading-UDTs#post5540423
  [udt_pass_clx]: https://vbforums.com/showthread.php?599355-RESOLVED-Addin-a-user-defined-type-to-a-collection
  [udt_pass_dix]: https://mrexcel.com/board/threads/is-it-possible-to-assign-udt-as-item-of-collection-dictionary.1221049#post-5971115
  [vba_typ_fn]:   https://learn.microsoft.com/office/vba/language/reference/user-interface-help/typename-function
  [vba_typ_op]:   https://learn.microsoft.com/dotnet/visual-basic/language-reference/operators/typeof-operator
  [sob_tmp_mtd]:  src/SObTemplate.bas#L217-L246
  [udt_inact]:    http://cpearson.com/excel/classes.aspx#:~:text=cannot%20carry%20out%20actions
  [vb_net]:       https://learn.microsoft.com/dotnet/visual-basic
  [vba_tostring]: https://stackoverflow.com/posts/comments/98934630
  [net_tostring]: https://learn.microsoft.com/dotnet/fundamentals/runtime-libraries/system-object-tostring
  [sob_secure]:   src/SOb.bas#L498-L513
  [obj_inst]:     http://cpearson.com/excel/classes.aspx#:~:text=New%20instances%20of%20a%20class%20may%20be%20created
  [vba_new]:      https://learn.microsoft.com/dotnet/visual-basic/language-reference/operators/new-operator
  [udt_inst]:     http://cpearson.com/excel/classes.aspx#:~:text=you%20can%27t%20declare%20new%20instances%20of%20a%20Type
  [obj_act]:      http://cpearson.com/excel/classes.aspx#:~:text=classes%20have%20methods
  [obj_valid]:    http://cpearson.com/excel/classes.aspx#:~:text=properties%20of%20a%20class%20can%20be%20set%20or%20retrieved
