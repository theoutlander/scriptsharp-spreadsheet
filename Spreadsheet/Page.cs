// Page.cs
//

using System;
using Spreadsheet;

[ScriptModule]
internal static class Page
{
    static Page()
    {
        // Create a new instance of the spreadsheet
        Sheet sheet = new Sheet();

        // Render the spreadsheet in the "spreadsheet" div element of the html file
        sheet.Render("spreadsheet");
    }
}