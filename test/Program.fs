open ClosedXML
open ClosedXML.Excel
open ClosedXML.SimpleSheets
open System.IO
open System

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

let createFullExample() : byte[] =
    use workbook = new XLWorkbook()
    let simpleFields = workbook.AddWorksheet("Simple Fields")
    Excel.populate(simpleFields, users, [
        Excel.field(fun user -> user.Name)
        Excel.field(fun user -> user.Age)
        Excel.field(fun user -> user.Working)
        Excel.field(fun user -> user.LastName)
        Excel.field(fun user -> user.DateOfBirth)
        Excel.field(fun user -> user.StartedJob)
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

    Excel.createFrom(workbook)

let simpleExample() = Excel.createFrom(users, [
    Excel.field(fun user -> user.Name)
    Excel.field(fun user -> user.Age)
])

File.WriteAllBytes("Simple.xlsx", simpleExample())
File.WriteAllBytes("FullExample.xlsx", createFullExample())