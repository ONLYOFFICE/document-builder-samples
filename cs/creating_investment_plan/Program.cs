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

using System.Text.Json;
using System.IO;

namespace Sample
{
    public class CreatingPresentation
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.xlsx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateInvestmentPlan(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateInvestmentPlan(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string jsonPath = resourcesDir + "/data/investment_data.json";
            string json = File.ReadAllText(jsonPath);
            InvestmentData data = JsonSerializer.Deserialize<InvestmentData>(json);

            // init docbuilder and create new xlsx file
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue worksheet = api.Call("GetActiveSheet");

            // initialize financial data from JSON
            int initAmount = data.initial_amount;
            double rate = data.return_rate;
            int term = data.term;

            // fill years
            CValue startCell = worksheet.Call("GetRangeByNumber", 1, 0);
            CValue endCell = worksheet.Call("GetRangeByNumber", term + 1, 0);
            string[] years = new string[term + 1];
            for (int year = 0; year <= term; year++)
            {
                years[year] = year.ToString();
            }
            worksheet.Call("GetRange", startCell, endCell).Call("SetValue", CreateColumnData(years));

            // fill initial amount
            worksheet.Call("GetRangeByNumber", 1, 1).Call("SetValue", initAmount);
            // fill remaining cells
            startCell = worksheet.Call("GetRangeByNumber", 2, 1);
            endCell = worksheet.Call("GetRangeByNumber", term + 1, 1);
            string[] amounts = new string[term];
            for (int year = 0; year < term; year++)
            {
                amounts[year] = $"=$B$2*POWER((1+{rate.ToString().Replace(',', '.')}),A{year + 3})";
            }
            worksheet.Call("GetRange", startCell, endCell).Call("SetValue", CreateColumnData(amounts));

            // create chart
            CValue chart = worksheet.Call("AddChart", $"Sheet1!$A$1:$B${term + 2}", false, "lineNormal", 2, 135.38 * 36000, 81.28 * 36000);
            chart.Call("SetPosition", 3, 0, 2, 0);
            chart.Call("SetTitle", "Capital Growth Over Time", 22);
            CValue color = api.Call("CreateRGBColor", 134, 134, 134);
            CValue fill = api.Call("CreateSolidFill", color);
            CValue stroke = api.Call("CreateStroke", 1, fill);
            chart.Call("SetMinorVerticalGridlines", stroke);
            chart.Call("SetMajorHorizontalGridlines", stroke);
            // fill table headers
            worksheet.Call("GetRangeByNumber", 0, 0).Call("SetValue", "Year");
            worksheet.Call("GetRangeByNumber", 0, 1).Call("SetValue", "Amount");

            // save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static CValue CreateColumnData(string[] data)
        {
            CValue[] columnData = new CValue[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                CValue[] row = new CValue[1];
                row[0] = data[i];
                columnData[i] = row;
            }
            return columnData;
        }
    }

    // Define classes to represent the JSON structure
    public class InvestmentData
    {
        public int initial_amount { get; set; }
        public double return_rate { get; set; }
        public int term { get; set; }
    }
}
