module Phonetic.Web

open System
open HtmlAgilityPack
open Phonetic.Config
open System.Linq

let combineUrl (word: string) =
    if appConf.WeblioUrl.ToString().Last().Equals('/') then
        $"{appConf.WeblioUrl}{word}"
    else
        $"{appConf.WeblioUrl}/{word}"

type Web() =
    member _.Read(urls: list<option<string>>) =
        seq {
            for oUrl in urls ->
                match oUrl with
                | Some url ->
                    try
                        let doc = (HtmlWeb()).Load(combineUrl url)

                        let phonetic =
                            doc
                                .DocumentNode
                                .SelectSingleNode(
                                    "//*[@id=\"phoneticEjjeNavi\"]/div/span[2]"
                                )
                                .InnerText
                            |> Some

                        let mean =
                            doc
                                .DocumentNode
                                .SelectSingleNode(
                                    "//*[@id=\"summary\"]/div[2]/p/span[2]"
                                )
                                .InnerText
                            |> Some

                        [ phonetic; mean ]
                    with
                    | :? NullReferenceException -> [ None; None ]
                    | x ->
                        printfn $"{url} をダウンロード中にエラーが発生しました\n{x}"
                        [ None; None ]
                | None -> [ None; None ]
        }
        |> Seq.toList
