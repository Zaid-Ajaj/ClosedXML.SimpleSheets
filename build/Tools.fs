﻿[<RequireQualifiedAccess>]
module Tools

open Nuke.Common.Tooling

let dotnet = ToolPathResolver.GetPathExecutable("dotnet")
let npm = ToolPathResolver.GetPathExecutable("npm")
let node = ToolPathResolver.GetPathExecutable("node")