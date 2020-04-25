# COM Type Provider

The COM Type Provider provides a new way to do COM interop from F#. Support consuming both 32-bit and 64-bit COM components. This repository is forked from [this fsprojects repository](https://github.com/fsprojects/FSharp.Interop.ComProvider). You can also find the [original documentation](http://fsprojects.github.io/FSharp.ComProvider/) here.

This repository is updated and maintenained by Wuping Xin (https://github.com/wxinix). 

## Technical Overview

Normally, to do COM interop from a .NET project, you need to use the _Add Reference_ context menu from the Project Explorer of the Visual Studio IDE and select the registered COM component you would like to reference. This generates an assembly containing the interop types that you can then consume from your code.

Behind the scenes, _Add Reference_ actually depends on the `TypeLibConverter` class to create the interop types in the generated assembly. As a matter of fact, the TypeLibConverter class is also used by the standalone `tlbimp.exe` tool. With this FSharp COM Type Provider, we leverage the same `TypeLibConverter` class  to generate the interop types.

## Limitations and Known Issues

The following known issues and limitations currently apply to the COM provider.
Some of them I would like to eventually rectify if possible:

* Type libraries with Primary Interop Assemblies (PIAs) such as Microsoft Office are not supported.
* All the types generated from the type library will be embedded in the assembly, rather than just the ones you actually refer to in your code.

## Maintainer of this Repository
- [@wxinix](https://github.com/wxinix)
