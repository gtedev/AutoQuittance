module Config

open System.Text.Json
open System.IO

type Template =
    { pathFile: string
      outputFile: string
      properties: Map<string, string> }

type Config = { templates: Template list }

let getConfig () =
    File.ReadAllText("Config.json") |> JsonSerializer.Deserialize<Config>
