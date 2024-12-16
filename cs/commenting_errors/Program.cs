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
    public class CommentingErrors
    {
        public static void Main(string[] args)
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.xlsx";
            string filePath = "../../../../../../resources/docs/spreadsheet_with_errors.xlsx";
            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CommentErrors(workDirectory, resultPath, filePath);
        }

        public static void CommentErrors(string workDirectory, string resultPath, string filePath)
        {
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;

            // Init DocBuilder
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();

            oBuilder.OpenFile(filePath, "xlsx");
            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];

            // Find and comment formula errors
            CValue oWorksheet = oApi.Call("GetActiveSheet");
            CValue oRange = oWorksheet.Call("GetUsedRange");
            var data = oRange.Call("GetValue");

            for (int row = 0; row < data.GetLength(); row++)
            {
                for (int col = 0; col < data[0].GetLength(); col++)
                {
                    CheckCell(oWorksheet, data[row][col].ToString(), row, col);
                }
            }

            // Save file and close DocBuilder
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void CheckCell(CValue oWorksheet, string cell, int row, int col)
        {
            if (cell.Contains("#"))
            {
                string comment = "Error" + cell;
                CValue errorCell = oWorksheet.Call("GetRangeByNumber", row, col);
                errorCell.Call("AddComment", comment);
            }
        }
    }
}
