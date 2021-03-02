open ClosedXML
open ClosedXML.Excel
open ClosedXML.SimpleSheets
open System.IO
open System
open Microcharts
open SkiaSharp
open ClosedXML.Excel.Drawings

type User = {
    Name: string;
    Age: int;
    Working: bool;
    LastName: string option
    DateOfBirth : DateTime
    StartedJob : DateTime option
}

let users = [
    { Name = "Zaid"; Age = 24; Working = true; LastName = Some "Ajaj"; DateOfBirth = DateTime(1996, 11, 13); StartedJob = None }
    { Name = "Jane"; Age = 26; Working = true; LastName = Some "Doe"; DateOfBirth = DateTime(2020, 01, 01); StartedJob = None }
    { Name = "John"; Age = 25; Working = false; LastName = None; DateOfBirth = DateTime(2020, 01, 01); StartedJob = Some(DateTime(2020, 01, 01)) }
]

type Website = {
    name: string
    address: string
}

[<RequireQualifiedAccess>]
type ChartType = BarChart | LineChart

type RowWithImage = {
    title: string
    data : float list
    chart: ChartType
}

let rnd = System.Random()

type Charts() =
    static member line(data: float seq) =
        let chart = LineChart()
        chart.Entries <- [
            for i in data -> ChartEntry(float32(i))
        ]

        chart.IsAnimated <- false
        chart.PointSize <- 0.0f
        chart.LineSize <- 1.0f
        chart.MinValue <- 0.0f
        chart.MaxValue <- 100.0f
        chart

    static member bar(data: float seq) =
        let chart = BarChart()
        chart.Entries <- [
            for i in data do
                let entry = ChartEntry(float32(i))
                entry
        ]

        chart.IsAnimated <- false
        chart.MinValue <- 0.0f
        chart.MaxValue <- 100.0f
        chart

    static member createImage(chart: Chart, width:int, height:int) =
        use bitmap = new SkiaSharp.SKBitmap(width, height)
        use canvas = new SkiaSharp.SKCanvas(bitmap)
        chart.DrawContent(canvas, width, height)
        canvas.Save() |> ignore
        use image = SKImage.FromPixels(bitmap.PeekPixels())
        use data = image.Encode(SKEncodedImageFormat.Png, 100)
        use memoryStream = new MemoryStream()
        data.AsStream().CopyTo(memoryStream)
        memoryStream.ToArray()

let rowWithImages = [
    {
        title =  "Line chart"
        data = [ for i in 1 .. 12 -> rnd.NextDouble() * 100.0 ]
        chart = ChartType.LineChart
    }

    {
        title = "Bar chart"
        data = [ for i in 1 .. 12 -> rnd.NextDouble() * 100.0 ]
        chart = ChartType.BarChart
    }
]

let createFullExample() : byte[] =
    use workbook = new XLWorkbook()
    let simpleFields = workbook.AddWorksheet("Simple Fields")
    Excel.populate(simpleFields, users, [
        Excel.field(fun user -> user.Name)
        Excel.field(fun user -> user.Age)
        Excel.field(fun user -> user.Working)
        Excel.field(fun user -> user.LastName)
        Excel.field(fun user -> user.DateOfBirth).adjustToContents()
        Excel.field(fun user -> user.StartedJob).adjustToContents()
    ])

    let fieldsWithHeaders = workbook.AddWorksheet("Added Headers")
    Excel.populate(fieldsWithHeaders, users, [
        Excel.field(fun user -> user.Name).header("Name")
        Excel.field(fun user -> user.Age).header("Age")
        Excel.field(fun user -> user.Working).header("Working")
    ])

    let styledRows = workbook.AddWorksheet("Styled Rows")
    Excel.populate(styledRows, users, [
        Excel.field(fun user -> user.Name)
            .header("Name")
            .strikethrough(fun user -> not user.Working)

        Excel.field(fun user -> user.Age)
            .header("Age")
            .headerFontColor(XLColor.White)
            .headerBackgroundColor(XLColor.DarkCyan)

        Excel.field(fun user -> user.Working)
            .header("Working")
            .headerFontColor(XLColor.White)
            .headerBackgroundColor(XLColor.DarkCyan)
            .fontColor(XLColor.White)
            .backgroundColor(fun user ->
                if user.Working
                then XLColor.Green
                else XLColor.Red
            )

        Excel.field(fun user -> user.DateOfBirth)
            .header("DateOfBirth")
            .headerFontColor(XLColor.White)
            .headerBackgroundColor(XLColor.DarkCyan)
            .adjustToContents()
            .dateFormat("dd/mm/yyyy")
    ])

    let sheetWithLinks = workbook.AddWorksheet("Adding Links")

    let websites = [
        { name = "Github"; address = "https://www.github.com" }
    ]

    Excel.populate(sheetWithLinks, websites, [
        Excel.field(fun website -> website.name)
            .hyperlink(fun website -> Uri(website.address))

        Excel.field(fun website -> Uri(website.address))
    ])

    let imageSheet = workbook.AddWorksheet("Images")
    Excel.populate(imageSheet, rowWithImages, [
        Excel.field(fun row -> row.title)
            .header("Title")
            .headerHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .verticalAlignment(XLAlignmentVerticalValues.Center)
            .horizontalAlignment(XLAlignmentHorizontalValues.Center)
            .columnWidth(15)

        Excel.field(fun row ->
            match row.chart with
            | ChartType.BarChart ->
                let barChart = Charts.bar(row.data)
                let chartBytes = Charts.createImage(chart=barChart, width=500, height=500)
                XLImage(chartBytes, XLPictureFormat.Png)

            | ChartType.LineChart ->
                let lineChart = Charts.line(row.data)
                let chartBytes = Charts.createImage(chart=lineChart, width=500, height=500)
                XLImage(chartBytes, XLPictureFormat.Png)
            )
            .header("Year overview")
            .headerHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .columnWidth(40)
            .rowHeight(60)
    ])

    Excel.createFrom(workbook)

let simpleExample() = Excel.createFrom(users, [
    Excel.field(fun user -> user.Name)
    Excel.field(fun user -> user.Age)
])

File.WriteAllBytes("Simple.xlsx", simpleExample())
File.WriteAllBytes("FullExample.xlsx", createFullExample())