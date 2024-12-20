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
    public class CommentingErrors
    {
        public static void Main()
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
            CDocBuilder builder = new();

            builder.OpenFile(filePath, "xlsx");
            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];

            // Find and comment formula errors
            CValue worksheet = api.Call("GetActiveSheet");
            CValue range = worksheet.Call("GetUsedRange");
            var data = range.Call("GetValue");

            for (int row = 0; row < data.GetLength(); row++)
            {
                for (int col = 0; col < data[0].GetLength(); col++)
                {
                    CheckCell(worksheet, data[row][col].ToString(), row, col);
                }
            }

            // Save file and close DocBuilder
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void CheckCell(CValue worksheet, string cell, int row, int col)
        {
            if (cell.Contains("#"))
            {
                string comment = "Error" + cell;
                CValue errorCell = worksheet.Call("GetRangeByNumber", row, col);
                errorCell.Call("AddComment", comment);
            }
        }
    }
}
