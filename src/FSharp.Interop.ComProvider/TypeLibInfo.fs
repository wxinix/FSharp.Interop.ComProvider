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

open Microsoft.Win32
open System
open System.IO
open Utility

type TypeLibVersion = {
    VersionStr: string;
    Major: int;
    Minor: int }

type TypeLib = {
    Name: string;
    Version: TypeLibVersion;
    Platform: string;
    Path: string
    Pia: string option } // Primary Interop Assembly

/// Extracts type lib version info from a string in the format major.minor, where major and minor are hex numbers, e.g, d.0.
let private tryParseVersion(text: string) =
    let hexToInt s = Int32.TryParse(s, Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture)
    match text.Split('.') |> Array.map hexToInt with
    | [| true, major; true, minor |] -> Some { VersionStr = text; Major = major; Minor = minor }
    | _ -> None

let private isInDotNetPath =
    let dotNetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"Microsoft.NET")
    fun (path: string) -> path.StartsWith(dotNetPath, StringComparison.OrdinalIgnoreCase)

/// Loads type libs for the preferred platform and returns a sequence of TypeLib records. The preferred platform
/// can be either "win32", "win64", or "*". If "*", then both "win32" and "win64" type libs will be loaded. 
let loadTypeLibs preferredPlatform (nameFilter: string) =
    [ use rootKey = Registry.ClassesRoot.OpenSubKey("TypeLib")
      for typeLibKey in rootKey.GetSubKeys() do                   // typeLibKey is a GUID string
          for verKey in typeLibKey.GetSubKeys() do                // verKey is a version string in the format "major.minor"
              for localeKey in verKey.GetSubKeys() do             // localeKey is a number representing locale
                  for platformKey in localeKey.GetSubKeys() do    // platformKey is either "win32" or "win64"
                      let name, version, locale = verKey.DefaultValue, tryParseVersion verKey.ShortName, localeKey.ShortName
                      if name <> "" && (name.ToLower().Contains(nameFilter.ToLower()) || nameFilter = "*") && version.IsSome && locale = "0" then                 
                          yield {
                              Name = name
                              Version = version.Value
                              Platform = platformKey.ShortName
                              Path = platformKey.DefaultValue
                              Pia = match verKey.GetValue("PrimaryInteropAssemblyName") with
                                    | :? string as pia -> Some pia
                                    | _ -> None } ]
    |> Seq.filter (fun lib -> not (isInDotNetPath lib.Path))
    |> Seq.groupBy (fun lib -> lib.Name, lib.Version)
    |> Seq.collect (fun (_, libs) -> libs |> Seq.filter(fun lib -> preferredPlatform = lib.Platform || preferredPlatform = "*"))