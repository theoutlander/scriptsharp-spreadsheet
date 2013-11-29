// Sheet.cs
//

using System;
using System.Html;

namespace Spreadsheet
{
    public class Sheet
    {
        private readonly TableElement _table;    // Table element that will contain the spreadsheet

        /// <summary>
        /// Spreadsheet Constructor
        /// </summary>
        public Sheet()
        {
            // Create a Table Element in memory
            _table = (TableElement)Document.CreateElement("table");

            // Render Column Headers
            RenderColumnHeaders();

            // Render Row Headers
            RenderRows(25);
        }

        /// <summary>
        /// Create rows in the spreadsheet
        /// </summary>
        /// <param name="rowCount">Number of rows to create</param>
        private void RenderRows(int rowCount)
        {
            // We're iterating from row index 1 because we want the title of the rows 
            // to be equal to the index. In addition, row 0 is the column header row
            for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                // Create a new row in the table
                TableRowElement row = _table.InsertRow(rowIndex);

                // Create a row header cell and set its id and value to the row index
                TableCellElement cell = row.InsertCell(0);
                cell.ID = rowIndex.ToString();
                cell.TextContent = rowIndex.ToString();

                // Create cells for each column in the spreadsheet from A to Z
                for (int cellIndex = 65; cellIndex < 91; cellIndex++)
                {
                    // Insert cells at the corresponding column index (starting from column 1)
                    cell = row.InsertCell(cellIndex - 64);

                    // Create a text input element inside the cell
                    InputElement input = (InputElement)Document.CreateElement("input");
                    input.Type = "text";

                    // Set the ID of the element to the Column Letter and Row Number like A1, B1, etc.
                    input.ID = string.FromCharCode(cellIndex) + rowIndex;

                    // Add the input element as a child of the cell
                    cell.AppendChild(input);

                    // Create and attach spreadsheet events to this input element
                    AttachEvents(input);
                }
            }
        }

        /// <summary>
        /// Create various events that tie the functionality of the spreadsheet
        /// </summary>
        /// <param name="input">Input element to attach events to</param>
        private void AttachEvents(InputElement input)
        {
            input.AddEventListener("blur", delegate(ElementEvent @event)
            {
                GetRowHeader(@event.SrcElement).ClassList.Remove("selected");
                GetColumnHeader(@event.SrcElement).ClassList.Remove("selected");

                ProcessCell((InputElement)@event.SrcElement);
            }, false);

            input.AddEventListener("focus", delegate(ElementEvent @event)
            {
                GetRowHeader(@event.SrcElement).ClassList.Add("selected");
                GetColumnHeader(@event.SrcElement).ClassList.Add("selected");

                object formula = @event.SrcElement.GetAttribute("data-formula");

                if (formula != null && formula.ToString().Length > 1)
                {
                    input.Value = formula.ToString();
                }

            }, false);

            input.AddEventListener("keydown", delegate(ElementEvent @event)
            {
                if (@event.KeyCode == 27) //Escape
                {
                    input.Value = "";
                    input.Blur();
                }
                else if (@event.KeyCode == 38) // up arrow
                {
                    SetFocusFromCellTo(@event.SrcElement.ID, 0, -1);
                    @event.PreventDefault();
                }
                else if (@event.KeyCode == 40) // down arrow
                {
                    SetFocusFromCellTo(@event.SrcElement.ID, 0, 1);
                    @event.PreventDefault();
                }
                else if (@event.KeyCode == 37) // left arrow
                {
                    SetFocusFromCellTo(@event.SrcElement.ID, -1, 0);
                    @event.PreventDefault();
                }
                else if (@event.KeyCode == 39) // right arrow
                {
                    SetFocusFromCellTo(@event.SrcElement.ID, 1, 0);
                    @event.PreventDefault();
                }
            }, true);

            input.AddEventListener("keypress", delegate(ElementEvent @event)
            {
                if (@event.KeyCode == 13) // Enter Key
                {
                    input.Blur();

                    SetFocusFromCellTo(@event.SrcElement.ID, 0, 1);
                    @event.PreventDefault();
                }
            }, false);
        }

        /// <summary>
        /// This method allows us to specify a cell and set focus to one of its neighboring cells
        /// </summary>
        /// <param name="id">Text Input Element id for the source cell</param>
        /// <param name="horizontal">Direction to move horizontally. Negative value moves left.</param>
        /// <param name="vertical">Direction to move vertically. Negative value moves up.</param>
        private void SetFocusFromCellTo(string id, int horizontal, int vertical)
        {
            int row = int.Parse(id.Substring(1));
            string column = id.Substring(0, 1);

            InputElement element = Document.GetElementById<InputElement>(string.Format("{0}{1}", string.FromCharCode(column.CharCodeAt(0) + horizontal), row + vertical));
            if (element != null)
            {
                element.Focus();
            }
        }

        private TableCellElement GetColumnHeader(Element element)
        {
            return Document.GetElementById<TableCellElement>(element.ID.Substring(0, 1));
        }

        private TableRowElement GetRowHeader(Element element)
        {
            return Document.GetElementById<TableRowElement>(element.ID.Substring(1));
        }

        private void RenderColumnHeaders()
        {
            // Create a row for the column headers
            TableRowElement header = _table.InsertRow();
            header.ID = "headerRow";

            // Create a cell for the first column where we will also add row headers
            TableCellElement blankCell = header.InsertCell(0);
            blankCell.ID = "blankCell";

            // Create 26 Column Headers and name them with the letters of the English Alphabets
            // Iterating through the ASCII indexes for A-Z
            for (int i = 65; i < 91; i++)
            {
                // Create a cell element in the header row starting at index 1 
                TableCellElement cell = header.InsertCell(i - 64);

                // Set the cell id to the letter corresponding to the ASCII index
                cell.ID = string.FromCharCode(i);

                // Set the value of the cell toe the letter corresponding to the ASCII index
                cell.TextContent = string.FromCharCode(i);
            }
        }

        /// <summary>
        /// Process the text entered in a text input element.
        /// Extract a formula and store it as part of the element.
        /// Process the formula and display the result in the text element.
        /// </summary>
        /// <param name="input">Input element of a cell to process</param>
        private void ProcessCell(InputElement input)
        {
            // Ensure that there's a value in the input element that is a formula involving another cell 
            if (input.Value.Length > 4 && input.Value.StartsWith("="))
            {
                // Set the input value as a data-formula attribute on the text input element
                input.SetAttribute("data-formula", input.Value);

                // For this tutorial, we will split the formula on the "+" operation only
                string[] items = input.Value.Substring(1).Split("+");

                Number result = 0;

                // Traverse through each item in the equation
                foreach (string item in items)
                {
                    // If the item is not a number, it is assumed to be a formula
                    if (Number.IsNaN((Number)(object)item))
                    {
                        // Get a reference to the cell, parse its value and then add it to our result
                        result += Number.Parse(((InputElement)Document.GetElementById(item)).Value);
                    }
                    else
                    {
                        result += Number.Parse(item);
                    }
                }

                // Replace the input's formula with the result. We've stored a reference to the formula as part of its data-formula attribute 
                input.Value = result.ToString();
            }
        }

        /// <summary>
        /// This function renders the spreadsheet in an HTML element
        /// </summary>
        /// <param name="divName">Name of a container div to create the spreadsheet in</param>
        public void Render(string divName)
        {
            // Get a reference to the Div Element and add the table as its child element
            Document.GetElementById<DivElement>(divName).AppendChild(_table);
        }
    }
}