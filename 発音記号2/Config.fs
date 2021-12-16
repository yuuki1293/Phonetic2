module Phonetic.Config

open FSharp.Data

type ConfigType = JsonProvider<"config.json">

let appConf = ConfigType.Parse "config.json"

type HtmlType = HtmlProvider<"https://ejje.weblio.jp/content/hello">

let html content = HtmlType.Parse content
