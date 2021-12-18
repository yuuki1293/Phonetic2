module Phonetic.Excel

open System.Diagnostics
open System.Text.RegularExpressions
open ClosedXML.Excel
open Microsoft.Win32
open Phonetic.Config
open Phonetic.Computation
open System.IO

let getExePath =
    if appConf.ExcelPath.ToString().Length > 0 then
        let rKeyName =
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Extensions"

        let rGetValueName = "xlsx"

        try
            let rKey =
                Registry.CurrentUser.OpenSubKey(rKeyName)

            let locate = rKey.GetValue rGetValueName |> string
            rKey.Close()
            Some locate
        with
        | x -> None
    else
        Some(appConf.ExcelPath.ToString())

let removeBracket str =
    let regex = Regex(@"\(.+?\)")
    regex.Replace(str, "")

let getWorksheet (workbook: XLWorkbook) =
    try
        Some(workbook.Worksheet(appConf.WorkSheetName.ToString()))
    with
    | _ -> None

let lastRow (workbook: XLWorkbook) =
    maybe {
        let! sheet = getWorksheet workbook
        return sheet.LastRowUsed().RowNumber()
    }

type Excel(inWorkbook: option<XLWorkbook>) =
    let workbook = inWorkbook
    let worksheet = inWorkbook >>= getWorksheet

    static member Open() =
        if File.Exists(appConf.FileName) then
            try
                Some(new XLWorkbook(appConf.FileName))
            with
            | x -> None
        else
            use file = new XLWorkbook()
            file.AddWorksheet("Sheet1") |> ignore
            file.SaveAs(appConf.FileName)
            Some file

    member _.Run() =
        maybe {
            let! path = getExePath

            let p =
                Process.Start(path, appConf.FileName.ToString())

            p.WaitForExit()
        }

    member _.Read() =
        (maybe {
            let! row = workbook >>= lastRow
            let! sheet = worksheet

            return
                (seq { for i in 1 .. row ->
                        match sheet.Cell(i, 1).Value.ToString() |> removeBracket with
                        | "" -> None
                        | x -> Some x
                     })
                |> Seq.toList
        }).Value

    member _.Write(phonetics: list<list<option<string>>>) =
        if worksheet.IsSome && workbook.IsSome then
            for i in 0 .. phonetics.Length - 1 do
                match phonetics[ i ][ 0 ] with
                | Some x -> $"[{x}]"
                | None -> "Not Found"
                |> worksheet.Value.Cell(i + 1, 2).SetValue
                |> ignore

                match phonetics[ i ][ 1 ] with
                | Some x -> x
                | None -> "Not found"
                |> worksheet.Value.Cell(i + 1, 3).SetValue

                |> ignore

            workbook.Value.Save()

    new() = Excel(Excel.Open())
    new(x: XLWorkbook) = Excel(Some x)
