
`ole-list` generates a (*VERY LARGE*) json file, containing information about available OLE typelibs, their classes, and their associated methods.  
It uses `win32ole` from the Ruby standard library. Obviously, this implies that `ole-list` will only work on some kind of NT-based platform (i.e., Windows itself, [Wine](http://winehq.org/), or [Reactos](http://reactos.org/)).

The json file will have this structure:

```json
    {
        "<the-name-of-this-typelib-which-will-usually-contain-spaces>":
        {
            "<the-name-of-this-class>":
            {
                "name": "<the-class-identifier>",
                "progid": "<the-ole-typename> (the string you'd use for CreateObject(), in case you're wondering!)", 
                "guid": "<{the-guid-for-this-class}>",
                "version": "0.0 (this is the version. shocking!)",
                "helpstring": "<some-kind-of-helpstring-or-null>"
                "methods":
                {
                    "<the-name-of-this-method>":
                    {
                        "return_type": "<some-builtin-ole-type>",
                        "invkind": "<type-of-invocation>",
                        "helpstring": "<some-helpstring-or-null>",
                        "proto": "<a-generated-pseudo-function-prototype>"
                        "params": [
                            {
                              "name": "<name-of-this-parameter>",
                              "type": "<some-builtin-ole-type>",
                              "is_input": <true|false>,
                              "is_output": <true|false>,
                              "is_optional": <true|false>
                            },
                            ...
                        ],
                    },
                    ... other methods ...
                },
            },
            ... other classes ...
        },
        ... other typelibs ...
    }
```

To get a better idea of how to traverse `json.list`, take a look at `mkpseudo.rb`, which generates the `pseudoheader.h`.

