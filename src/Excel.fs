namespace ClosedXML.SimpleSheets

open ClosedXML
open ClosedXML.Excel
open System
open System.IO

type FieldMap<'T> =
    {
        CellTransformers : ('T -> IXLCell -> IXLCell) list
        HeaderTransformers : (IXLCell -> IXLCell) list
    }
    with
        static member empty<'T>() = { CellTransformers = []; HeaderTransformers = [] }
        static member create<'T>(mapRow: 'T -> IXLCell -> IXLCell) =
            let empty = FieldMap<'T>.empty()
            { empty with CellTransformers = List.append empty.CellTransformers [mapRow] }

        member self.header(name: string) =
            let transformer (cell: IXLCell) = cell.SetValue(name)
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.headerBackgroundColor(color: XLColor) =
            let transformer (cell: IXLCell) =
                cell.Style.Fill.BackgroundColor <- color
                cell
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.headerFontColor(color: XLColor) =
            let transformer (cell: IXLCell) =
                cell.Style.Font.FontColor <- color
                cell
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.headerFontSize(size: float) =
            let transformer (cell: IXLCell) =
                cell.Style.Font.FontSize <- size
                cell
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.backgroundColor(color: XLColor) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Fill.BackgroundColor <- color
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.backgroundColor(color: 'T -> XLColor) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Fill.BackgroundColor <- color row
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.fontColor(color: XLColor) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.FontColor <- color
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.fontColor(color: 'T -> XLColor) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.FontColor <- color row
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.fontSize(size: float) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.FontSize <- size
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.fontSize(size: 'T -> float) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.FontSize <- size row
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.bold() =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Bold <- true
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.bold(bold: bool) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Bold <- bold
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.bold(bold: 'T -> bool) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Bold <- bold row
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.italic() =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Italic <- true
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.italic(italic: bool) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Italic <- italic
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.italic(italic: 'T -> bool) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Bold <- italic row
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.strikethrough() =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Strikethrough <- true
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.strikethrough(strikethrough: bool) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Strikethrough <- strikethrough
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.strikethrough(strikethrough: 'T -> bool) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Strikethrough <- strikethrough row
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.underline(underline: XLFontUnderlineValues) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Underline <- underline
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.underline(underline: 'T -> XLFontUnderlineValues) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Font.Underline <- underline row
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.dateFormat(format: string) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.DateFormat.Format <- format
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.numberFormat(format: string) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.NumberFormat.Format <- format
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.transformCell(transform: 'T -> IXLCell -> IXLCell) =
            { self with CellTransformers = List.append self.CellTransformers [transform] }

type Excel() =
    static member field<'T>(map: 'T -> int) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> string) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> DateTime) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> bool) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> double) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> int option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> DateTime option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> bool option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> double option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> string option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | None -> cell.SetValue(null)
        | Some text -> cell.SetValue(text)
    )
    static member populate<'T>(sheet: IXLWorksheet, data: seq<'T>, fields: FieldMap<'T> list) : unit =
        let headerTransformerGroups = fields |> List.map (fun field -> field.HeaderTransformers)
        let noHeadersAvailable =
            headerTransformerGroups
            |> List.concat
            |> List.isEmpty

        let headersAvailable = not noHeadersAvailable

        if headersAvailable then
            for (headerIndex, headerTransformers) in List.indexed headerTransformerGroups do
                let activeHeaderCell = sheet.Row(1).Cell(headerIndex + 1)
                for header in headerTransformers do ignore (header activeHeaderCell)

        for (rowIndex, row) in Seq.indexed data do
            let startRowIndex = if headersAvailable then 2 else 1
            let activeRow = sheet.Row(rowIndex + startRowIndex)
            for (fieldIndex, field) in List.indexed fields do
                let activeCell = activeRow.Cell(fieldIndex + 1)
                for transformer in field.CellTransformers do
                    ignore (transformer row activeCell)

    static member workbookToBytes(workbook: XLWorkbook) =
        use memoryStream = new MemoryStream()
        workbook.SaveAs(memoryStream)
        memoryStream.ToArray()

    static member createFrom(name: string, data: seq<'T>, fields: FieldMap<'T> list) : byte[] =
        use workbook = new XLWorkbook()
        let sheet = workbook.AddWorksheet(name)
        Excel.populate(sheet, data, fields)
        Excel.workbookToBytes(workbook)

    static member createFrom(workbook: XLWorkbook) =
        use memoryStream = new MemoryStream()
        workbook.SaveAs(memoryStream)
        memoryStream.ToArray()

    static member createFrom(data: seq<'T>, fields: FieldMap<'T> list) : byte[] =
        Excel.createFrom("Sheet1", data, fields)

    static member contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"