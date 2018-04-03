module Main
open System
open FSharp.Data
open System.Linq
open ClosedXML.Excel

[<Literal>]
let InputSrc = __SOURCE_DIRECTORY__ + @"\..\..\invoice-sent_2017_on_2018-3-24.csv"

type Csv = CsvProvider< InputSrc, ";",PreferOptionals=true>

type InvoiceOutputDto = {
    RecordNo: int;
    DatumVpisa: DateTime;
    StKnjigovodskeListine:string;
    DatumKnjListine: DateTime;
    OpisPoslDogogka: string;
    Kupec: string;
    ZnesekListine: decimal;
    ZnesekPrihodkovPoSRSAliZdoh2: decimal;
    ZnesekPopravkaDohodkov: decimal;
    DatumPlacila: DateTime;
    Opombe: string;
}

let mapToDto (idx,invoice:Csv.Row) =
    {
        RecordNo = idx;
        DatumVpisa = invoice.Date_sent;
        StKnjigovodskeListine = invoice.Docnum;
        DatumKnjListine = invoice.Date_sent;
        OpisPoslDogogka = "Izdan račun za delo po pogodbi";
        Kupec = invoice.Contact_fullname;
        ZnesekListine = invoice.Amount_w_vat;
        ZnesekPrihodkovPoSRSAliZdoh2 = invoice.Amount_w_vat;
        ZnesekPopravkaDohodkov = 0M;
        DatumPlacila = invoice.Payed_dates;
        Opombe = ""
    }

let formatAsDate (x:IXLCell) =
        x.Style.DateFormat.Format <- "dd.MM.yyyy"
        x
let formatAsCurrency (x:IXLCell) =
        x.Style.NumberFormat.NumberFormatId <- 4
        x
let bold (x:IXLCell) =
        x.Style.Font.Bold <- true
        x

let createRow (ws:IXLWorksheet) (rowIdx, invoice: InvoiceOutputDto) =
    printfn "%d, %s" invoice.RecordNo invoice.StKnjigovodskeListine
    [   ws.Cell(rowIdx, 1).SetValue(invoice.RecordNo);
        ws.Cell(rowIdx, 2).SetValue(invoice.DatumVpisa)|>formatAsDate;
        ws.Cell(rowIdx, 3).SetValue(invoice.StKnjigovodskeListine);
        ws.Cell(rowIdx, 4).SetValue(invoice.DatumKnjListine)|>formatAsDate;
        ws.Cell(rowIdx, 5).SetValue(invoice.OpisPoslDogogka);
        ws.Cell(rowIdx, 6).SetValue(invoice.Kupec);
        ws.Cell(rowIdx, 7).SetValue(invoice.ZnesekListine)|>formatAsCurrency;
        ws.Cell(rowIdx, 8).SetValue(invoice.ZnesekPrihodkovPoSRSAliZdoh2)|>formatAsCurrency;
        ws.Cell(rowIdx, 9).SetValue(invoice.ZnesekPopravkaDohodkov)|>formatAsCurrency;
        ws.Cell(rowIdx, 10).SetValue(invoice.DatumPlacila)|>formatAsDate
    ]

[<EntryPoint>]
let main argv =
    printfn "%A" argv
    let a = Csv.Load(InputSrc).Rows
    let years = a|> Seq.map (fun x -> x.Date_served.Year) |> Seq.distinct
    let count = Seq.length a
    printfn "Got %d invoices. Years (as served): %s" (count) (String.Join(", ", years))
    let recordNumbers = Enumerable.Range(1, (count))
    let outputRecords = a|> Seq.zip recordNumbers|> Seq.map mapToDto
    printfn "Creating xlsx."
    let wb = new XLWorkbook()
    let ws = wb.Worksheets.Add("Sheet1")
    ws.Cell(1,1).SetValue("Informacijske rešitve, Jernej Logar s.p.") |> ignore
    ws.Cell(2,1).SetValue("Novi svet 14, 4220 Škofja Loka") |> ignore
    ws.Cell(3,1).SetValue("Davčna številka: 71910778") |> ignore
    ws.Cell(5,1).SetValue("Zap. št. vpisa") |> ignore
    ws.Cell(5,2).SetValue("Datum vpisa") |> ignore
    ws.Cell(5,3).SetValue("Št. knjigovodske listine") |> ignore
    ws.Cell(5,4).SetValue("Datum knj. listine") |> ignore
    ws.Cell(5,5).SetValue("Opis posl. dogodka") |> ignore
    ws.Cell(5,6).SetValue("Kupec, posl. Partner") |> ignore
    ws.Cell(5,7).SetValue("Znesek listine") |> ignore
    ws.Cell(5,8).SetValue("Znesek prihodkov po SRS ali Zdoh-2") |> ignore
    ws.Cell(5,9).SetValue("Znesek popravka dohodkov") |> ignore
    ws.Cell(5,10).SetValue("Datum plačila") |> ignore
    ws.Cell(5,11).SetValue("Opombe") |> ignore
    let initialRow = 6;
    let createRow1 = (createRow ws)
    let cells = outputRecords
                    |> Seq.zip (Seq.map (fun x -> x - 1 + initialRow) recordNumbers)
                    |> Seq.map createRow1
    let sumRowIndex = initialRow + count
    ws.Cell(sumRowIndex,1).SetValue("SKUPAJ") |> ignore
    ws.Cell(sumRowIndex,1).Style.Font.Bold <- true

    let sumRows cellIndex =
        let c1 = Seq.tryHead cells |> Option.map (fun c -> (c.Item cellIndex).Address)
        let c2 = Seq.tryLast cells |> Option.map (fun c -> (c.Item cellIndex).Address)
        let sum first last =
            ws.Cell(sumRowIndex, cellIndex + 1) |>
                fun c ->
                    c.FormulaA1 <- (sprintf "SUM(%A:%A)" first last)
                    formatAsCurrency c |> bold |> ignore
                    c
        Option.map2 sum c1 c2
    
    [6;7;8]
        |> List.map sumRows
        |> ignore
    Enumerable.Range(2, 11)
        |> Seq.toList
        |> List.map (fun idx -> ws.Column(idx).AdjustToContents())
        |> ignore 
    wb.SaveAs("c:\\temp\\out.xlsx")
    printfn "ok."
    Console.ReadLine()|>ignore
    0 // return an integer exit code
