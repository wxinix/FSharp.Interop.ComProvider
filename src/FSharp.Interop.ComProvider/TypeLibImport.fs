﻿// MIT License
// Copyright (c) Wuping Xin 2020.
//
// Permission is hereby  granted, free of charge, to any  person obtaining a copy
// of this software and associated  documentation files (the "Software"), to deal
// in the Software  without restriction, including without  limitation the rights
// to  use, copy,  modify, merge,  publish, distribute,  sublicense, and/or  sell
// copies  of  the Software,  and  to  permit persons  to  whom  the Software  is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE  IS PROVIDED "AS  IS", WITHOUT WARRANTY  OF ANY KIND,  EXPRESS OR
// IMPLIED,  INCLUDING BUT  NOT  LIMITED TO  THE  WARRANTIES OF  MERCHANTABILITY,
// FITNESS FOR  A PARTICULAR PURPOSE AND  NONINFRINGEMENT. IN NO EVENT  SHALL THE
// AUTHORS  OR COPYRIGHT  HOLDERS  BE  LIABLE FOR  ANY  CLAIM,  DAMAGES OR  OTHER
// LIABILITY, WHETHER IN AN ACTION OF  CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE  OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

// The UNLICENSE
// Copyright (c) Microsoft Corporation 2005-2012.
// A license with no conditions whatsoever  which  dedicates works to the  public
// domain. Unlicensed  works, modifications,  and larger works may be distributed
// under different terms and without source code.

module private FSharp.Interop.ComProvider.TypeLibImport

open System
open System.IO
open System.Reflection
open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes
open ReflectionProxies
open TypeLibDoc

[<DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)>]
extern void private LoadTypeLib(string filename, ITypeLib& typelib);

let loadTypeLib path =
    let mutable typeLib:ITypeLib = null
    LoadTypeLib(path, &typeLib)
    if typeLib = null then
        failwith ("Error loading type library. Please check that the component is correctly installed and registered.")
    typeLib

let hideEvents assembly =
    // COM events don't import correctly as first-class events in F#, as it
    // expects the first parameter to be the "sender", a convention COM does
    // not follow. Therefore, we hide the actual event members and only expose
    // the add / remove handler methods.
    let hideTypeEvents ty =
        { new TypeProxy(ty) with
            override __.GetEvents(flags) = [||] } :> Type
    { new AssemblyProxy(assembly) with
        override __.GetTypes() = base.GetTypes() |> Array.map hideTypeEvents } :> Assembly

let rec importTypeLib path asmDir =
    let assemblies = ResizeArray<Assembly>()
    let rec convertToAsm (typeLib:ITypeLib) =
        let converter = TypeLibConverter()
        let libName = Marshal.GetTypeLibName(typeLib)
        let asmFile = libName + ".dll"
        let asmPath = Path.Combine(asmDir, asmFile)
        let flags = TypeLibImporterFlags.None
        let sink = { new ITypeLibImporterNotifySink with
            member __.ReportEvent(eventKind, eventCode, eventMsg) = ()
            member __.ResolveRef(typeLib) = convertToAsm (typeLib :?> ITypeLib) }
        let asm = converter.ConvertTypeLibToAssembly(typeLib, asmPath, flags, sink, null, null, libName, null)
        asm.Save(asmFile)
        let typeDocs = getTypeLibDoc typeLib
        Assembly.LoadFrom(asmPath)
        |> annotateAssembly typeDocs
        |> hideEvents
        |> assemblies.Add
        asm :> Assembly
    let typeLib = loadTypeLib path
    convertToAsm typeLib |> ignore
    assemblies |> Seq.toList