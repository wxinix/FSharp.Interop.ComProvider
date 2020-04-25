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

module private FSharp.Interop.ComProvider.TypeLibInfo

open System
open System.IO
open Microsoft.Win32
open Utility

type TypeLibVersion = {
    String:string;
    Major:int;
    Minor:int }

type TypeLib = {
    Name:string;
    Version:TypeLibVersion;
    Platform:string;
    Path:string
    Pia:string option }

let private tryParseVersion (text:string) =
    match text.Split('.') |> Array.map Int32.TryParse with
    | [|true, major; true, minor|] -> Some { String = text; Major = major; Minor = minor }
    | _ -> None

let private isInDotNetPath =
    let dotNetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"Microsoft.NET")
    fun (path:string) -> path.StartsWith(dotNetPath, StringComparison.OrdinalIgnoreCase)

let loadTypeLibs preferredPlatform =
    [ use rootKey = Registry.ClassesRoot.OpenSubKey("TypeLib")
      for typeLibKey in rootKey.GetSubKeys() do
      for versionKey in typeLibKey.GetSubKeys() do
      for localeKey in versionKey.GetSubKeys() do
      for platformKey in localeKey.GetSubKeys() do
          let name = versionKey.DefaultValue
          let version = tryParseVersion versionKey.SubKeyName
          if name <> ""
             && version.IsSome
             && localeKey.SubKeyName = "0"
          then
              yield { Name = name
                      Version = version.Value
                      Platform = platformKey.SubKeyName
                      Path = platformKey.DefaultValue
                      Pia = match versionKey.GetValue("PrimaryInteropAssemblyName") with
                            | :? string as pia -> Some pia
                            | _ -> None } ]
    |> Seq.filter (fun lib -> not (isInDotNetPath lib.Path))
    |> Seq.groupBy (fun lib -> lib.Name, lib.Version)
    |> Seq.map (fun (_, libs) ->
        match libs |> Seq.tryFind (fun lib -> lib.Platform = preferredPlatform) with
        | Some lib -> lib
        | None -> Seq.head libs)
