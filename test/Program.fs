open ClosedXML
open ClosedXML.Excel
open ClosedXML.SimpleSheets
open System.IO

type User = { Name: string; Age: int; Active: bool }

let users = [
    { Name = "Zaid"; Age = 24; Active = true }
    { Name = "Jane"; Age = 26; Active = false }
    { Name = "John"; Age = 25; Active = false }
]

let createFullExample() : byte[] =
    use workbook = new XLWorkbook()
    let simpleFields = workbook.AddWorksheet("Simple Fields")
    Excel.populate(simpleFields, users, [
        Excel.field(fun user -> user.Name)
        Excel.field(fun user -> user.Age)
        Excel.field(fun user -> user.Active)
    ])

    let fieldsWithHeaders = workbook.AddWorksheet("Added Headers")
    Excel.populate(fieldsWithHeaders, users, [
        Excel.field(fun user -> user.Name).header("Name")
        Excel.field(fun user -> user.Age).header("Age")
        Excel.field(fun user -> user.Active).header("Active")
    ])

    let styledRows = workbook.AddWorksheet("Styled Rows")
    Excel.populate(styledRows, users, [
        Excel.field(fun user -> user.Name)
            .header("Name")
            .strikethrough(fun user -> not user.Active)

        Excel.field(fun user -> user.Age)
            .header("Age")
            .headerFontColor(XLColor.White)
            .headerBackgroundColor(XLColor.DarkCyan)

        Excel.field(fun user -> user.Active)
            .header("Active")
            .fontColor(XLColor.White)
            .backgroundColor(fun user ->
                if user.Active
                then XLColor.Green
                else XLColor.Red
            )
    ])

    Excel.createFrom(workbook)

let simpleExample() = Excel.createFrom(users, [
    Excel.field(fun user -> user.Name)
    Excel.field(fun user -> user.Age)
])

File.WriteAllBytes("Simple.xlsx", simpleExample())
File.WriteAllBytes("FullExample.xlsx", createFullExample())