namespace ClosedXML.SimpleSheets

open System
open System.IO

open ClosedXML
open ClosedXML.Excel
open ClosedXML.Excel.Drawings

type XLImage(content: byte[], format: XLPictureFormat) =
    member self.content = content
    member self.format = format

    new (content: byte[]) = XLImage(content, XLPictureFormat.Png)



type FieldMap<'T> =
    {
        CellTransformers : ('T -> IXLCell -> IXLCell) list
        HeaderTransformers : (IXLCell -> IXLCell) list
        ColumnWidth : float option
        RowHeight : ('T -> float option) option
        AdjustToContents: bool
    }
    with
        static member empty<'T>() = {
            CellTransformers = []
            HeaderTransformers = []
            ColumnWidth = None
            RowHeight = None
            AdjustToContents = false
        }

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

        member self.hyperlink(link: 'T -> Uri) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.SetHyperlink(XLHyperlink(link row))
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.hyperlink(link: 'T -> Uri option) =
            let transformer (row: 'T) (cell: IXLCell) =
                match link row with
                | Some uri ->
                    cell.SetHyperlink(XLHyperlink(uri))
                    cell
                | None ->
                    cell

            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.hyperlink(link: 'T -> XLHyperlink) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.SetHyperlink(link row)
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.hyperlink(link: 'T -> XLHyperlink option) =
            let transformer (row: 'T) (cell: IXLCell) =
                match link row with
                | Some hyperlink ->
                    cell.SetHyperlink(hyperlink)
                    cell
                | None ->
                    cell

            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.horizontalAlignment(alignment: XLAlignmentHorizontalValues) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Alignment.Horizontal <- alignment
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.verticalAlignment(alignment: XLAlignmentVerticalValues) =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Alignment.Vertical <- alignment
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.centered() =
            let transformer (row: 'T) (cell: IXLCell) =
                cell.Style.Alignment.Vertical <- XLAlignmentVerticalValues.Center
                cell.Style.Alignment.Horizontal <- XLAlignmentHorizontalValues.Center
                cell
            { self with CellTransformers = List.append self.CellTransformers [transformer] }

        member self.headerHorizontalAlignment(alignment: XLAlignmentHorizontalValues) =
            let transformer (cell: IXLCell) =
                cell.Style.Alignment.Horizontal <- alignment
                cell
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.headerVerticalAlignment(alignment: XLAlignmentVerticalValues) =
            let transformer (cell: IXLCell) =
                cell.Style.Alignment.Vertical <- alignment
                cell
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.headerCentered() =
            let transformer (cell: IXLCell) =
                cell.Style.Alignment.Vertical <- XLAlignmentVerticalValues.Center
                cell.Style.Alignment.Horizontal <- XLAlignmentHorizontalValues.Center
                cell
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.columnWidth(width: float) =
            { self with ColumnWidth = Some width }

        member self.columnWidth(width: int) =
            { self with ColumnWidth = Some (float width) }

        member self.columnWidth(width: float option) =
            { self with ColumnWidth = width }

        member self.rowHeight(height: int) =
            { self with RowHeight = Some(fun row -> Some (float height)) }

        member self.rowHeight(height: float) =
            { self with RowHeight = Some(fun row -> Some height) }

        member self.rowHeight(height: float option) =
            { self with RowHeight = Some(fun row -> height) }

        member self.rowHeight(height: 'T -> float option) =
            { self with RowHeight = Some(fun row -> height row) }

        member self.adjustToContents() =
            { self with AdjustToContents = true }

        member self.transformCell(transform: 'T -> IXLCell -> IXLCell) =
            { self with CellTransformers = List.append self.CellTransformers [transform] }

type Excel() =
    static member field<'T>(map: 'T -> int) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> string) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> DateTime) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> bool) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> double) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> int option) = FieldMap<'T>.create(fun row cell -> 
        let value= Option.toNullable (map row)
        cell.SetValue(value))
    static member field<'T>(map: 'T -> DateTime option) = FieldMap<'T>.create(fun row cell -> 
        let value = Option.toNullable (map row)
        cell.SetValue(value))
    static member field<'T>(map: 'T -> bool option) = FieldMap<'T>.create(fun row cell -> 
        let value = Option.toNullable (map row) 
        cell.SetValue(XLCellValue.FromObject(value)))
    static member field<'T>(map: 'T -> double option) = FieldMap<'T>.create(fun row cell -> 
        let value = Option.toNullable (map row)
        cell.SetValue(value))
    static member field<'T>(map: 'T -> string option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | None -> cell
        | Some text -> cell.SetValue(text)
    )

    static member field<'T>(map: 'T -> DateTimeOffset) = FieldMap<'T>.create(fun row cell ->
        let value = map row
        cell.SetValue(value.UtcDateTime)
    )

    static member field<'T>(map: 'T -> DateTimeOffset option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | None -> cell
        | Some value -> cell.SetValue(value.UtcDateTime)
    )

    static member field<'T>(map: 'T -> Uri) = FieldMap<'T>.create(fun row cell ->
        let uri = map row
        cell.SetHyperlink(XLHyperlink(uri))
        cell.SetValue(uri.ToString())
    )

    static member field<'T>(map: 'T -> Uri option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | Some uri ->
            cell.SetHyperlink(XLHyperlink(uri))
            cell.SetValue(uri.ToString())
        | None ->
            cell
    )

    static member field<'T>(map: 'T -> decimal) = FieldMap<'T>.create(fun row cell ->
        let value = Convert.ToDouble(map row)
        cell.SetValue(value)
    )

    static member field<'T>(map: 'T -> decimal option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | None -> cell
        | Some value -> cell.SetValue(Convert.ToDouble(value))
    )

    static member field<'T>(map: 'T -> int64) = FieldMap<'T>.create(fun row cell ->
        let value = Convert.ToDouble(map row)
        cell.SetValue(value)
    )

    static member field<'T>(map: 'T -> int64 option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | None -> cell
        | Some value -> 
            let valueAsDouble = Convert.ToDouble(value)
            cell.SetValue(valueAsDouble)
    )

    static member field<'T>(map: 'T -> Guid) = FieldMap<'T>.create(fun row cell ->
        let value = map row
        cell.SetValue(value.ToString())
    )

    static member field<'T>(map: 'T -> Guid option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | None -> cell
        | Some value -> cell.SetValue(value.ToString())
    )

    static member field<'T>(map: 'T -> XLImage) = FieldMap<'T>.create(fun row cell ->
        let image = map row
        let worksheet = cell.Worksheet
        let addedImage = worksheet.AddPicture(new MemoryStream(image.content), image.format)
        addedImage.MoveTo(cell, cell.CellBelow().CellRight()) |> ignore
        addedImage.Placement <- XLPicturePlacement.MoveAndSize
        cell
    )

    static member field<'T>(map: 'T -> XLImage option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | Some image ->
            let worksheet = cell.Worksheet
            let addedImage = worksheet.AddPicture(new MemoryStream(image.content), image.format)
            addedImage.MoveTo(cell, cell.CellBelow().CellRight()) |> ignore
            addedImage.Placement <- XLPicturePlacement.MoveAndSize
            cell
        | None ->
            cell
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

                if field.AdjustToContents then
                    let currentColumn = activeCell.WorksheetColumn()
                    currentColumn.AdjustToContents() |> ignore
                    activeRow.AdjustToContents() |> ignore

                match field.ColumnWidth with
                | Some givenWidth ->
                    let currentColumn = activeCell.WorksheetColumn()
                    currentColumn.Width <- givenWidth
                | None -> ()

                match field.RowHeight with
                | Some givenHeightFn ->
                    match givenHeightFn row with
                    | Some givenHeight ->
                        activeRow.Height <- givenHeight
                    | None ->
                        ()
                | None ->
                    ()

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

[<AutoOpen>]
module Extensions =
    type System.Collections.Generic.IEnumerable<'T> with
        member inline data.excelField(map: 'T -> string) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> string option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> bool) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> bool option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> int) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> int option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> double) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> double option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> decimal) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> decimal option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> DateTime) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> DateTime option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> DateTimeOffset) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> DateTimeOffset option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> int64) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> int64 option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> Guid) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> Guid option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> Uri) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> Uri option) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> XLImage) : FieldMap<'T> = Excel.field(map)
        member inline data.excelField(map: 'T -> XLImage option) : FieldMap<'T> = Excel.field(map)