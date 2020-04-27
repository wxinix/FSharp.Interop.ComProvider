# COM Type Provider

The COM Type Provider provides a new way to do COM interop from F#. Support consuming both 32-bit and 64-bit COM components. This repository is cloned from the original [fsprojects/FSharp.Interop.ComProvider](https://github.com/fsprojects/FSharp.Interop.ComProvider) project. You can also find the [original documentation](http://fsprojects.github.io/FSharp.ComProvider/) here.

This cloned repository is seperately updated and maintenained by Wuping Xin (https://github.com/wxinix). It serves as my personal "dogfooding" project to learn F# language.

## Technical Overview

Normally, to do COM interop from a .NET project, you need to use the _Add Reference_ context menu from the Project Explorer of the Visual Studio IDE and select the registered COM component you would like to reference. This generates an assembly containing the interop types that you can then consume from your code.

Behind the scenes, _Add Reference_ actually depends on the `TypeLibConverter` class to create the interop types in the generated assembly. As a matter of fact, the TypeLibConverter class is also used by the standalone `tlbimp.exe` tool. With this FSharp COM Type Provider, we leverage the same `TypeLibConverter` class  to generate the interop types.

## Enhancements and Changes

The original [fsprojects/FSharp.Interop.ComProvider](https://github.com/fsprojects/FSharp.Interop.ComProvider) project has not been updated since 2016 when F# 3.0 was still in Beta.

With this "dogbooding" project of my own, I made the following updates and changes to the orignal [fsprojects/FSharp.Interop.ComProvider](https://github.com/fsprojects/FSharp.Interop.ComProvider) project:
- Bring it up to date with F# 4.7.1.
- Remove the limitation that only 32 bit COM components are supported. Now both 32 bit and 64 bit can be consumed.
- Cleanse the project layout, removing those files that are irrelevant for my learning purpose.

## Limitations and Known Issues

The following known issues and limitations currently apply to the COM provider.
Some of them I would like to eventually rectify if possible:

* Type libraries with Primary Interop Assemblies (PIAs) such as Microsoft Office are not supported.
* All the types generated from the type library will be embedded in the assembly, rather than just the ones you actually refer to in your code.

## Code Style Guidline

Microsoft [coding style guideline]https://docs.microsoft.com/en-us/dotnet/fsharp/style-guide/formatting is used.

## Maintainer of this Repository
- [@wxinix](https://github.com/wxinix)
