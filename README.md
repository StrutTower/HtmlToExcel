# TowerSoft HtmlToExcel

Small Nuget package that uses AngleSharp to read an HTML table and generate an Excel file using ClosedXML

### Usage

Single sheet
```csharp
string htmlString = "<table><tbody><tr><td>Cell contents</td></tr></tbody></table>";

byte[] fileData = new WorkbookGenerator().FromHtmlString(htmlString);
```

Multiple sheets
```csharp
string htmlString1 = "<table><tbody><tr><td>Cell contents</td></tr></tbody></table>";
string htmlString2 = "<table><tbody><tr><td>Cell contents</td></tr></tbody></table>";

byte[] fileData;
using (WorkbookBuilder workbookBuilder = new WorkbookBuilder()) {
    workbookBuilder.AddSheet("sheet1", htmlString1);
    workbookBuilder.AddSheet("sheet2", htmlString2);

    fileData = workbookBuilder.GetAsByteArray();
}
```

### Settings

Some settings do not work if there are any colspans in the table.
This is because Excel does not allow tables with merged cells and those settings only work in tables.

| Setting Name | Default Value | Description |
|--------------|---------------|-------------|
| AutofitColumns | true | Enables/disables fitting the width of the columns to fit the contents. |
| ShowFilter | true | Enables/disables showing table filters. Does not work with the table has any colspans. |
| ShowRowStripes | true | Enables/disables row stripes. Does not work with the table has any colspans. |

You can change the settings using `HtmlToExcellSettings`
and passing it in the constructor of `htmlToEzxel`

```csharp
HtmlToExcelSettings settings = HtmlToExcelSettings.Defaults;
settings.AutofitColumns = false;
settings.ShowFilter = false;
settings.ShowRowStripes = false;

// Using custom settings with a single sheet
byte[] fileData new WorkbookGenerator(settings).FromHtmlString(htmlString);

// Using custom settings with multiple sheets
using (WorkbookBuilder workbookBuilder = new WorkbookBuilder(settings) {
    // Settings can also be used in the AddSheet method which overrides the setting on the WorkbookBuilder
    workbookBuilder.AddSheet("sheetName", htmlString, settings)
}
```

#### Individual Cell Options


| Attribute Name | Expected Data Type | Comments |
|----------------|--------------------|----------|
| data-excel-hyperlink | URI | Creates a hyperlink on the cell. Must be a parsable absolute URI. |
| data-excel-bold | Boolean | Sets if the cell style will be set to bold. |
| data-type | String | Sets the data type for the cell. Valid options are: Text, Number, Boolean, DateTime, TimeSpan |
| data-format | String | Sets the cell format. Only works with data-type="Number" and data-type="DateTime" |
| data-excel-comment | String | Adds a comment to the cell |
| data-excel-comment-author | String | Sets the author for the comment |
| data-horizontal-alignment | String | Sets the cell horizontal alignment. Valid options are: Center, CenterContinuous, Distributed, Fill, General, Justify, Left, Right |
| colspan | Integer | Merges this cell with the following cells |
| data-font-size | Integer | Sets the cell font size |
| data-font-color | String(HexColor) | Sets the font color of the cell. Must be a valid hex color code, Example: #006688 |
| data-background-color | String(HexColor) | Sets the fill/background color of the cell. Must be a valid hex color code, Example: #006688 |


#### ASP Core Example
Add the following code to your project to render a view to a string:
[CustomController.cs](https://gist.github.com/StrutTower/da303d31f2c930cb5a34af7a0968a0d3)

Use this example to return the file to the client. Make sure you change the inherited class to your custom controller class.

```csharp
public class HomeController : CustomController {
    public IActionResult ExcelFile() {
        var model = //Get model data

        string htmlString = RenderViewAsync("viewName", model, true);
        byte[] fileData = new WorkbookGenerator().FromHtmlString(htmlString);

        return File(fileData, MimeType.xlsx, "filename.xlsx");
    }
}
```


#### MVC 5 Example
Add the following code to your project to render a view to a string:
[ViewExtensions.cs](https://gist.github.com/StrutTower/d5aa7677f5bb22fb5a5c28c0faab885c)

```csharp
public ActionResult GetExcelFile() {
    var model = //Get model data

    string htmlString = PartialView("ViewName", model).RenderToString();
    byte[] fileData = new WorkbookGenerator().FromHtmlString(htmlString);

    return File(fileData, MimeType.xlsx, "filename.xlsx");
}
```
