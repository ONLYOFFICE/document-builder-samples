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
    public class FillingSpreadsheet
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.xlsx";
            object[,] data = {
                { "Id", "Product", "Price", "Available"},
                { 1001, "Item A", 12.2, true },
                { 1002, "Item B", 18.8, true },
                { 1003, "Item C", 70.1, false },
                { 1004, "Item D", 60.6, true },
                { 1005, "Item E", 32.6, true },
                { 1006, "Item F", 28.3, false },
                { 1007, "Item G", 11.1, false },
                { 1008, "Item H", 41.4, true }
            };
        // add Docbuilder dlls in path
        System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            FillSpreadsheet(workDirectory, resultPath, data);
        }

        public static void FillSpreadsheet(string workDirectory, string resultPath, object[,] data)
        {
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;

            // Init DocBuilder
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();

            builder.CreateFile(doctype);
            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue worksheet = api.Call("GetActiveSheet");

            // pass data
            CValue array = TwoDimArrayToCValue(data, context);
            // First cell in the range (A1) is equal to (0,0)
            CValue startCell = worksheet.Call("GetRangeByNumber", 0, 0);
            // Last cell in the range is equal to array length -1
            CValue endCell = worksheet.Call("GetRangeByNumber", array.GetLength() - 1, array[0].GetLength() - 1);
            worksheet.Call("GetRange", startCell, endCell).Call("SetValue", array);


            // Save file and close DocBuilder
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static CValue TwoDimArrayToCValue(object[,] data, CContext context)
        {
            int rowsLen = data.GetLength(0);
            int colsLen = data.GetLength(1);
            CValue array = context.CreateArray(rowsLen);

            for (int row = 0; row < rowsLen; row++)
            {
                CValue arrayCol = context.CreateArray(colsLen);

                for (int col = 0; col < colsLen; col++)
                {
                    arrayCol[col] = data[row, col].ToString();
                }
                array[row] = arrayCol;
            }
            return array;
        }
    }
}
