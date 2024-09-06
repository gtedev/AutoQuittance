module Config

open System.Text.Json
open System.IO

type Period =
    { startDayInMonth: int
      sendDayInMonth: int }

type Template =
    { pathFile: string
      outputFile: string
      period: Period
      displayComputedProperties: string list
      displayCustomProperties: Map<string, string> }

type Config = { templates: Template list }

let getConfig () =
    File.ReadAllText("Config.json") |> JsonSerializer.Deserialize<Config>
