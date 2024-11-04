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
using CContextScope = docbuilder_net.CDocBuilderContextScope;

namespace Sample
{
    public class CreatingChartPresentation
    {
        public static void Main(string[] args)
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
            CDocBuilder oBuilder = new CDocBuilder();

            // Read chart data from xlsx
            oBuilder.OpenFile(filePath, "xlsx");
            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oWorksheet = oApi.Call("GetActiveSheet");
            CValue oRange = oWorksheet.Call("GetUsedRange").Call("GetValue");
            object[,] array = oRangeTo2dArray(oRange, oContext);
            oBuilder.CloseFile();

            // Create chart presentation
            oBuilder.CreateFile(doctype);
            oContext = oBuilder.GetContext();
            oScope = oContext.CreateScope();
            oGlobal = oContext.GetGlobal();
            oApi = oGlobal["Api"];
            CValue oPresentation = oApi.Call("GetPresentation");
            CValue oSlide = oPresentation.Call("GetSlideByIndex", 0);
            oSlide.Call("RemoveAllObjects");

            CValue oRGBColor = oApi.Call("CreateRGBColor", 255, 244, 240);
            CValue oFill = oApi.Call("CreateSolidFill", oRGBColor);
            oSlide.Call("SetBackground", oFill);

            CValue oStroke = oApi.Call("CreateStroke", 0, oApi.Call("CreateNoFill"));
            CValue oShapeTitle = oApi.Call("CreateShape", "rect", 300 * 36000, 20 * 36000, oApi.Call("CreateNoFill"), oStroke);
            CValue oShapeText = oApi.Call("CreateShape", "rect", 120 * 36000, 80 * 36000, oApi.Call("CreateNoFill"), oStroke);
            oShapeTitle.Call("SetPosition", 20 * 36000, 20 * 36000);
            oShapeText.Call("SetPosition", 210 * 36000, 50 * 36000);
            CValue oParagraphTitle = oShapeTitle.Call("GetDocContent").Call("GetElement", 0);
            CValue oParagraphText = oShapeText.Call("GetDocContent").Call("GetElement", 0);
            oRGBColor = oApi.Call("CreateRGBColor", 115, 81, 68);
            oFill = oApi.Call("CreateSolidFill", oRGBColor);

            string titleContent = "Price Type Report";
            string textContent = "This is an overview of price types. As we can see, May was the price peak, but even in June the price went down, the annual upward trend persists.";
            addText(oApi, 80, titleContent, oSlide, oShapeTitle, oParagraphTitle, oFill, "center");
            addText(oApi, 42, textContent, oSlide, oShapeText, oParagraphText, oFill, "left");

            // Transform 2d array into cols names, rows names and data
            CValue array_cols = colsFromArray(array, oContext);
            CValue array_rows = rowsFromArray(array, oContext);
            CValue array_data = dataFromArray(array, oContext);

            // Pass CValue data to the CreateChart method
            CValue oChart = oApi.Call("CreateChart", "lineStacked", array_data, array_cols, array_rows);
            oChart.Call("SetSize", 180 * 36000, 100 * 36000);
            oChart.Call("SetPosition", 20 * 36000, 50 * 36000);
            oChart.Call("ApplyChartStyle", 24);
            oChart.Call("SetLegendFontSize", 12);
            oChart.Call("SetLegendPos", "top");
            oSlide.Call("AddObject", oChart);

            // Save file and close DocBuilder
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static object[,] oRangeTo2dArray(CValue oRange, CContext oContext)
        {
            int rowsLen = (int)oRange.GetLength();
            int colsLen = (int)oRange[0].GetLength();
            object[,] oArray = new object[rowsLen, colsLen];

            for (int col = 0; col < colsLen; col++)
            {
                CValue oArrayRow = oContext.CreateArray(rowsLen);
                for (int row = 0; row < rowsLen; row++)
                {
                    oArray[row, col] = oRange[row][col].ToString();
                }
            }
            return oArray;
        }

        public static CValue colsFromArray(object[,] array, CContext oContext)
        {
            int colsLen = array.GetLength(1) - 1;
            CValue cols = oContext.CreateArray(colsLen);
            for (int col = 1; col <= colsLen; col++)
            {
                cols[col - 1] = array[0, col].ToString();
            }
            return cols;
        }

        public static CValue rowsFromArray(object[,] array, CContext oContext)
        {
            int rowsLen = array.GetLength(0) - 1;
            CValue rows = oContext.CreateArray(rowsLen);
            for (int row = 1; row <= rowsLen; row++)
            {
                rows[row - 1] = array[row, 0].ToString();
            }
            return rows;
        }

        public static CValue dataFromArray(object[,] array, CContext oContext)
        {
            int colsLen = array.GetLength(0) - 1;
            int rowsLen = array.GetLength(1) - 1;
            CValue data = oContext.CreateArray(rowsLen);
            for (int row = 1; row <= rowsLen; row++)
            {
                CValue row_data = oContext.CreateArray(colsLen);
                for (int col = 1; col <= colsLen; col++)
                {
                    row_data[col - 1] = array[col, row].ToString();
                }
                data[row - 1] = row_data;
            }
            return data;
        }

        public static void addText(CValue oApi, int fontSize, string text, CValue oSlide, CValue oShape, CValue oParagraph, CValue oFill, string jc)
        {
            CValue oRun = oApi.Call("CreateRun");
            var oTextPr = oRun.Call("GetTextPr");
            oTextPr.Call("SetFontSize", fontSize);
            oTextPr.Call("SetFill", oFill);
            oTextPr.Call("SetFontFamily", "Tahoma");
            oParagraph.Call("SetJc", jc);
            oRun.Call("AddText", text);
            oRun.Call("AddLineBreak");
            oParagraph.Call("AddElement", oRun);
            oSlide.Call("AddObject", oShape);
        }
    }
}
