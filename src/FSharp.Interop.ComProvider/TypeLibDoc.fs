// MIT License
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

module private FSharp.Interop.ComProvider.TypeLibDoc

open Microsoft.FSharp.Core.CompilerServices
open System
open System.Collections.Generic
open System.Reflection
open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes
open ReflectionProxies
open Utility

let getStruct<'t when 't: struct> ptr freePtr =
    let st = Marshal.PtrToStructure(ptr, typeof<'t>) :?> 't
    freePtr ptr
    st

let getTypeLibDoc (typeLib: ITypeLib) =
    [ for typeIndex = 0 to typeLib.GetTypeInfoCount() - 1 do
        let typeInfo = typeLib.GetTypeInfo(typeIndex)
        let typeAttr = getStruct<TYPEATTR> (typeInfo.GetTypeAttr()) typeInfo.ReleaseTypeAttr
        let typeName, typeDoc, _, _ = typeInfo.GetDocumentation(-1)
        let memberDocs =
            [ for funcIndex = 0 to int typeAttr.cFuncs - 1 do
                let funcDesc = getStruct<FUNCDESC> (typeInfo.GetFuncDesc(funcIndex)) typeInfo.ReleaseFuncDesc
                let funcName, funcDoc, _, _ = typeInfo.GetDocumentation(funcDesc.memid)
                yield funcName, funcDoc ]
        yield typeName, (typeDoc, Map.ofSeq memberDocs) ]
    |> Map.ofSeq

let annotateAssembly typeDocs (asm: Assembly) =
    let toList (items: seq<'t>) = ResizeArray<'t> items :> IList<'t>

    let attrCons = typeof<TypeProviderXmlDocAttribute>.GetConstructor [| typeof<string> |]
    
    let attrData docString =
        { new CustomAttributeData() with
            override __.Constructor = attrCons
            override __.ConstructorArguments = [ CustomAttributeTypedArgument docString ] |> toList }
    
    let addAttr docString (memb: MemberInfo) =
        match docString with
        | s when not (String.IsNullOrWhiteSpace s) -> [attrData s]
        | _ -> []
        |> Seq.append (memb.GetCustomAttributesData())
        |> toList

    let findSourceInterface (ty: Type) =
        match ty.TryGetAttribute<ComEventInterfaceAttribute>() with
        | Some attr -> attr.SourceInterface
        | _ -> ty

    let findRelatedMember (memb: MemberInfo) =
        memb.DeclaringType.GetEvents()
        |> Seq.tryFind (fun event ->
            [ event.GetAddMethod(); event.GetRemoveMethod() ]
            |> Seq.exists(fun m -> m :> MemberInfo = memb))
        |> function
            | Some event -> event :> MemberInfo
            | None -> memb

    let typeDoc (ty: Type) =
        typeDocs
        |> Map.tryFind ty.Name
        |> Option.map fst

    let memberDoc (memb: MemberInfo) =
        let ty = memb.DeclaringType
        ty.GetInterfaces()
        |> Seq.append [ty]
        |> Seq.map findSourceInterface
        |> Seq.choose (fun ty -> typeDocs |> Map.tryFind ty.Name)
        |> Seq.choose (fun (_, membs) -> membs |> Map.tryFind (findRelatedMember memb).Name)
        |> Seq.tryFind (not << String.IsNullOrEmpty)

    let annotate getDoc addAnnotation (memb: #MemberInfo) =  // Object -> MemberInfo -> Type
        let doc = defaultArg (getDoc memb) ""
        addAnnotation (addAttr doc memb) memb

    let annotateEvent = annotate memberDoc <| fun data event ->
         { new EventInfoProxy(event) with
            override __.GetCustomAttributesData() = data } :> EventInfo

    let annotateMethod = annotate memberDoc <| fun data meth ->
        { new MethodInfoProxy(meth) with
            override __.GetCustomAttributesData() = data } :> MethodInfo

    let annotateProperty = annotate memberDoc <| fun data prop ->
        { new PropertyInfoProxy(prop) with
            override __.GetCustomAttributesData() = data } :> PropertyInfo

    let annotateType = annotate typeDoc <| fun attr ty ->  // Pipe backward
        { new TypeProxy(ty) with
            override __.GetCustomAttributesData() = attr            
            
            override __.GetEvents(flags) =
                ty.GetEvents(flags) |> Array.map annotateEvent            
            
            override __.GetMethods(flags) =
                ty.GetMethods(flags) |> Array.map annotateMethod
            
            override __.GetProperties(flags) =
                ty.GetProperties(flags) |> Array.map annotateProperty } :> Type

    { new AssemblyProxy(asm) with
        override __.GetTypes() = asm.GetTypes() |> Array.map annotateType } :> Assembly
