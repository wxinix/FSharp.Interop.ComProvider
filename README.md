# COM Type Provider

> *"An F# type provider is a component that provides types, properties, and methods for use in your program.  Type Providers generate what are known as Provided Types, which are generated by the F# compiler and are based on an external data source".* (see [Microsoft Document](https://docs.microsoft.com/en-us/dotnet/fsharp/tutorials/type-providers/)).  

Simply put, an F# type provider is just an F# compiler plugin, allowing a developer to add new types to the language's type system. The added types then become automatically visible to user code at compile time. 

COM Type Provider is a specialized F# type provider using registered COM type libraries as the data source, while providing a new way to do COM interop from F#.

## Purpose

Adapted from the [open-source fsprojects](https://github.com/fsprojects/FSharp.Interop.ComProvider), this repo is independently updated and maintenained by [Wuping Xin](https://github.com/wxinix), serving as a personal "dogfooding" project in F# language. You are welcome to reference and use it in your own project should you find it useful.

## Technical Overview

Normally, to do COM interop from a .NET project, you need to use the _Add Reference_ context menu from the Project Explorer of the Visual Studio IDE and select the registered COM component you would like to reference. This generates an assembly containing the interop types that you can then consume from your code.

Behind the scenes, _Add Reference_ actually depends on the `TypeLibConverter` class to create the interop types in the generated assembly. As a matter of fact, the TypeLibConverter class is also used by the standalone `tlbimp.exe` tool. With this FSharp COM Type Provider, we leverage the same `TypeLibConverter` class  to generate the interop types. With type provider, there is no need to explicitly import and reference the generated assembly. The "Provided Types" will automatically become part of the compiler's type system.

## Enhancements and Changes

The original [fsprojects/FSharp.Interop.ComProvider](https://github.com/fsprojects/FSharp.Interop.ComProvider) project has not been updated since 2016 when F# 3.0 was still in Beta.

The following updates, enhancements, and changes have been made to the orignal [fsprojects/FSharp.Interop.ComProvider](https://github.com/fsprojects/FSharp.Interop.ComProvider) project:
- Bring it up to date with F# 4.7.2.
- Remove the limitation that only 32 bit COM components are supported. Now both 32 bit and 64 bit can be consumed.
- Cleanse the project layout, removing those files that are irrelevant.
- Add a wildcard "*" as an additional option for preferred platform, intended for loading both win32 and win64 platforms.
- Enhance the namespace and type path, following the new but more meaningful format: COM.TypeLibraryName.Version-Platform.
- Various code cleanup and adding relevant comments.

## Limitations and Known Issues

The following known issues and limitations currently apply to the COM provider.
Some of them I would like to eventually rectify if possible:

* Type libraries with Primary Interop Assemblies (PIAs) such as Microsoft Office are not supported.
* All the types generated from the type library will be embedded in the assembly, regardless of whether they are actually referenced in the code or not.

## Code Style Guidline

Microsoft [coding style guideline for F#](https://docs.microsoft.com/en-us/dotnet/fsharp/style-guide/formatting) is used.

## Maintainer of this Repository
- [@wxinix](https://github.com/wxinix)  Wuping Xin
