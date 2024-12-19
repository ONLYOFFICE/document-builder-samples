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

using System.Collections.Generic;
using System.Text.Json;
using System.IO;

namespace Sample
{
    public class CreatingPresentation
    {
        public static void Main(string[] args)
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.xlsx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateAnnualReport(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateAnnualReport(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string json_path = resourcesDir + "/data/investment_data.json";
            string json = File.ReadAllText(json_path);
            InvestmentData data = JsonSerializer.Deserialize<InvestmentData>(json);

            // init docbuilder and create new xlsx file
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.CreateFile(doctype);

            CContext oContext = oBuilder.GetContext();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oWorksheet = oApi.Call("GetActiveSheet");

            // initialize financial data from JSON
            int initAmount = data.initial_amount;
            double rate = data.return_rate;
            int term = data.term;

            // fill years
            CValue oStartCell = oWorksheet.Call("GetRangeByNumber", 1, 0);
            CValue oEndCell = oWorksheet.Call("GetRangeByNumber", term + 1, 0);
            string[] years = new string[term + 1];
            for (int year = 0; year <= term; year++)
            {
                years[year] = year.ToString();
            }
            oWorksheet.Call("GetRange", oStartCell, oEndCell).Call("SetValue", createColumnData(years));

            // fill initial amount
            oWorksheet.Call("GetRangeByNumber", 1, 1).Call("SetValue", initAmount);
            // fill remaining cells
            oStartCell = oWorksheet.Call("GetRangeByNumber", 2, 1);
            oEndCell = oWorksheet.Call("GetRangeByNumber", term + 1, 1);
            string[] amounts = new string[term];
            for (int year = 0; year < term; year++)
            {
                amounts[year] = $"=$B$2*POWER((1+{rate.ToString().Replace(',', '.')}),A{year + 3})";
            }
            oWorksheet.Call("GetRange", oStartCell, oEndCell).Call("SetValue", createColumnData(amounts));

            // create chart
            CValue oChart = oWorksheet.Call("AddChart", $"Sheet1!$A$1:$B${term + 2}", false, "lineNormal", 2, 135.38 * 36000, 81.28 * 36000);
            oChart.Call("SetPosition", 3, 0, 2, 0);
            oChart.Call("SetTitle", "Capital Growth Over Time", 22);
            CValue oColor = oApi.Call("CreateRGBColor", 134, 134, 134);
            CValue oFill = oApi.Call("CreateSolidFill", oColor);
            CValue oStroke = oApi.Call("CreateStroke", 1, oFill);
            oChart.Call("SetMinorVerticalGridlines", oStroke);
            oChart.Call("SetMajorHorizontalGridlines", oStroke);
            // fill table headers
            oWorksheet.Call("GetRangeByNumber", 0, 0).Call("SetValue", "Year");
            oWorksheet.Call("GetRangeByNumber", 0, 1).Call("SetValue", "Amount");

            // save and close
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static CValue createColumnData(string[] data)
        {
            CValue[] arrColumnData = new CValue[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                CValue[] arrRow = new CValue[1];
                arrRow[0] = data[i];
                arrColumnData[i] = arrRow;
            }
            return arrColumnData;
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
