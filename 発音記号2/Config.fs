module Phonetic.Config

open System.IO
open FSharp.Data
open Phonetic

type ConfigType = JsonProvider<"config.json">

let createDefaultConfig=
    use stream = File.CreateText("config.json")
    stream.Write(
        """
        {
          "fileName" : "発音記号2.xlsx",
          "workSheetName" : "Sheet1",
          "excelPath" : "",
          "weblioUrl": "https://ejje.weblio.jp/content/"
        }
        """)

let appConf =
    if not (File.Exists "config.json") then createDefaultConfig
    try
        use stream = File.OpenRead("config.json")
        ConfigType.Load stream
    with
    | x ->
        printfn $"コンフィグがロードできませんでした\n"
        ConfigType.Load """
        {
          "fileName" : "発音記号2.xlsx",
          "workSheetName" : "Sheet1",
          "excelPath" : "",
          "weblioUrl": "https://ejje.weblio.jp/content/"
        }
        """