module Phonetic.Excel

open System.Diagnostics
open System.Text.RegularExpressions
open ClosedXML.Excel
open Microsoft.Win32
open Phonetic.Config
open Phonetic.Computation
open System.Linq

let getExePath =
    if appConf.ExcelPath.ToString().Length >0 then
        let rKeyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Extensions"
        let rGetValueName = "xlsx"
        try
            let rKey = Registry.CurrentUser.OpenSubKey(rKeyName)
            let locate = rKey.GetValue rGetValueName |> string
            rKey.Close()
            Some locate
        with
        | x ->
            None
    else
        Some (appConf.ExcelPath.ToString())

let removeBracket str=
    let regex = Regex(@"\(.+?\)");
    regex.Replace(str, "");

let getWorksheet(workbook:XLWorkbook) =
    try
        Some (workbook.Worksheet(appConf.WorkSheetName.ToString()))
    with
    | _ -> None

let lastRow (workbook:XLWorkbook) =
    maybe{
        let! sheet = getWorksheet workbook
        return sheet.LastRowUsed().RowNumber()
    }

type Excel(inWorkbook:option<XLWorkbook>,inWorksheet:Option<IXLWorksheet>) =
    let workbook = inWorkbook
    let worksheet = inWorkbook >>= getWorksheet
    member _.Open() =
        try
            Some (new XLWorkbook(appConf.FileName))
        with
        | x -> None
    member _.Run()=
        maybe {
            let! path = getExePath
            let p = Process.Start(path, appConf.FileName.ToString());
            p.WaitForExit()
        }
    member _.Read() =
        maybe{
            let! row = workbook >>= lastRow
            let! sheet = worksheet
            return (seq{
                for i in 1 .. row -> sheet.Cell(i,1).Value.ToString() |> removeBracket
            }).ToList()
        }
    new () = Excel (None,None)