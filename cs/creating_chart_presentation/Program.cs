/**
 *
 * (c) Copyright Ascensio System SIA 2024
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

using docbuilder_net;

using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;

namespace Sample
{
    public class CreatingChartPresentation
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.pptx";
            string filePath = "../../../../../../resources/docs/chart_data.xlsx";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateChartPresentation(workDirectory, resultPath, filePath);
        }

        public static void CreateChartPresentation(string workDirectory, string resultPath, string filePath)
        {
            var doctype = (int)OfficeFileTypes.Presentation.PPTX;

            // Init DocBuilder
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();

            // Read chart data from xlsx
            builder.OpenFile(filePath, "xlsx");
            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue worksheet = api.Call("GetActiveSheet");
            CValue range = worksheet.Call("GetUsedRange").Call("GetValue");
            object[,] array = RangeTo2dArray(range);
            builder.CloseFile();

            // Create chart presentation
            builder.CreateFile(doctype);
            context = builder.GetContext();
            global = context.GetGlobal();
            api = global["Api"];
            CValue presentation = api.Call("GetPresentation");
            CValue slide = presentation.Call("GetSlideByIndex", 0);
            slide.Call("RemoveAllObjects");

            CValue rGBColor = api.Call("CreateRGBColor", 255, 244, 240);
            CValue fill = api.Call("CreateSolidFill", rGBColor);
            slide.Call("SetBackground", fill);

            CValue stroke = api.Call("CreateStroke", 0, api.Call("CreateNoFill"));
            CValue shapeTitle = api.Call("CreateShape", "rect", 300 * 36000, 20 * 36000, api.Call("CreateNoFill"), stroke);
            CValue shapeText = api.Call("CreateShape", "rect", 120 * 36000, 80 * 36000, api.Call("CreateNoFill"), stroke);
            shapeTitle.Call("SetPosition", 20 * 36000, 20 * 36000);
            shapeText.Call("SetPosition", 210 * 36000, 50 * 36000);
            CValue paragraphTitle = shapeTitle.Call("GetDocContent").Call("GetElement", 0);
            CValue paragraphText = shapeText.Call("GetDocContent").Call("GetElement", 0);
            rGBColor = api.Call("CreateRGBColor", 115, 81, 68);
            fill = api.Call("CreateSolidFill", rGBColor);

            string titleContent = "Price Type Report";
            string textContent = "This is an overview of price types. As we can see, May was the price peak, but even in June the price went down, the annual upward trend persists.";
            AddText(api, 80, titleContent, slide, shapeTitle, paragraphTitle, fill, "center");
            AddText(api, 42, textContent, slide, shapeText, paragraphText, fill, "left");

            // Transform 2d array into cols names, rows names and data
            CValue array_cols = ColsFromArray(array, context);
            CValue array_rows = RowsFromArray(array, context);
            CValue array_data = DataFromArray(array, context);

            // Pass CValue data to the CreateChart method
            CValue chart = api.Call("CreateChart", "lineStacked", array_data, array_cols, array_rows);
            chart.Call("SetSize", 180 * 36000, 100 * 36000);
            chart.Call("SetPosition", 20 * 36000, 50 * 36000);
            chart.Call("ApplyChartStyle", 24);
            chart.Call("SetLegendFontSize", 12);
            chart.Call("SetLegendPos", "top");
            slide.Call("AddObject", chart);

            // Save file and close DocBuilder
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static object[,] RangeTo2dArray(CValue range)
        {
            int rowsLen = (int)range.GetLength();
            int colsLen = (int)range[0].GetLength();
            object[,] array = new object[rowsLen, colsLen];

            for (int col = 0; col < colsLen; col++)
            {
                for (int row = 0; row < rowsLen; row++)
                {
                    array[row, col] = range[row][col].ToString();
                }
            }
            return array;
        }

        public static CValue ColsFromArray(object[,] array, CContext context)
        {
            int colsLen = array.GetLength(1) - 1;
            CValue cols = context.CreateArray(colsLen);
            for (int col = 1; col <= colsLen; col++)
            {
                cols[col - 1] = array[0, col].ToString();
            }
            return cols;
        }

        public static CValue RowsFromArray(object[,] array, CContext context)
        {
            int rowsLen = array.GetLength(0) - 1;
            CValue rows = context.CreateArray(rowsLen);
            for (int row = 1; row <= rowsLen; row++)
            {
                rows[row - 1] = array[row, 0].ToString();
            }
            return rows;
        }

        public static CValue DataFromArray(object[,] array, CContext context)
        {
            int colsLen = array.GetLength(0) - 1;
            int rowsLen = array.GetLength(1) - 1;
            CValue data = context.CreateArray(rowsLen);
            for (int row = 1; row <= rowsLen; row++)
            {
                CValue row_data = context.CreateArray(colsLen);
                for (int col = 1; col <= colsLen; col++)
                {
                    row_data[col - 1] = array[col, row].ToString();
                }
                data[row - 1] = row_data;
            }
            return data;
        }

        public static void AddText(CValue api, int fontSize, string text, CValue slide, CValue shape, CValue paragraph, CValue fill, string jc)
        {
            CValue run = api.Call("CreateRun");
            var textPr = run.Call("GetTextPr");
            textPr.Call("SetFontSize", fontSize);
            textPr.Call("SetFill", fill);
            textPr.Call("SetFontFamily", "Tahoma");
            paragraph.Call("SetJc", jc);
            run.Call("AddText", text);
            run.Call("AddLineBreak");
            paragraph.Call("AddElement", run);
            slide.Call("AddObject", shape);
        }
    }
}
