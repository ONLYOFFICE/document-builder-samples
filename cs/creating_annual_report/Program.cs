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
            string resultPath = "../../../result.docx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateAnnualReport(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateAnnualReport(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string json_path = resourcesDir + "/data/financial_system_response.json";
            string json = File.ReadAllText(json_path);
            YearData data = JsonSerializer.Deserialize<YearData>(json);

            // init docbuilder and create new docx file
            var doctype = (int)OfficeFileTypes.Document.DOCX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.CreateFile(doctype);

            CContext oContext = oBuilder.GetContext();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oDocument = oApi.Call("GetDocument");

            // DOCUMENT HEADER
            CValue oParagraph = oDocument.Call("GetElement", 0);
            addTextToParagraph(oParagraph, $"Annual Report for {data.year}", 44, true, "center");

            // FINANCIAL section
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Financial performance", 32, true);
            oDocument.Call("Push", oParagraph);
            // quarterly data
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Quarterly data:", 24);
            oDocument.Call("Push", oParagraph);
            // chart
            oParagraph = oApi.Call("CreateParagraph");
            string[] chartKeys = { "revenue", "expenses", "net_profit" };
            var quarterlyData = data.financials.quarterly_data;
            CValue[] arrChartData = new CValue[chartKeys.Length];
            for (int i = 0; i < chartKeys.Length; i++)
            {
                arrChartData[i] = oContext.CreateArray(quarterlyData.Count);
                for (int j = 0; j < quarterlyData.Count; j++)
                {
                    arrChartData[i][j] = (int)quarterlyData[j].GetType().GetProperty(chartKeys[i]).GetValue(quarterlyData[j]);
                }
            }
            CValue[] arrChartNames = { "Revenue", "Expenses", "Net Profit" };
            CValue[] arrHorValues = { "Q1", "Q2", "Q3", "Q4" };
            CValue oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, arrChartNames, arrHorValues);
            oChart.Call("SetSize", 170 * 36000, 90 * 36000);
            oParagraph.Call("AddDrawing", oChart);
            oDocument.Call("Push", oParagraph);
            // expenses
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Expenses:", 24);
            oDocument.Call("Push", oParagraph);
            // pie chart
            oParagraph = oApi.Call("CreateParagraph");
            int rdExpenses = data.financials.r_d_expenses;
            int marketingExpenses = data.financials.marketing_expenses;
            int totalExpenses = data.financials.total_expenses;
            arrChartData = new CValue[1];
            arrChartData[0] = new CValue[] { rdExpenses, marketingExpenses, totalExpenses - (rdExpenses + marketingExpenses) };
            arrChartNames = new CValue[] { "Research and Development", "Marketing", "Other" };
            oChart = oApi.Call("CreateChart", "pie", arrChartData, oContext.CreateArray(0), arrChartNames);
            oChart.Call("SetSize", 170 * 36000, 90 * 36000);
            oParagraph.Call("AddDrawing", oChart);
            oDocument.Call("Push", oParagraph);
            // year totals
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Year total numbers:", 24);
            oDocument.Call("Push", oParagraph);
            // table
            CValue oTable = createTable(oApi, 2, 3);
            fillTableHeaders(oTable, new string[] { "Total revenue", "Total expenses", "Total net profit" }, 22);
            oParagraph = getTableCellParagraph(oTable, 1, 0);
            addTextToParagraph(oParagraph, $"{data.financials.total_revenue}", 22);
            oParagraph = getTableCellParagraph(oTable, 1, 1);
            addTextToParagraph(oParagraph, $"{data.financials.total_expenses}", 22);
            oParagraph = getTableCellParagraph(oTable, 1, 2);
            addTextToParagraph(oParagraph, $"{data.financials.net_profit}", 22);
            oDocument.Call("Push", oTable);

            // ACHIEVEMENTS section
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Achievements this year", 32, true);
            oDocument.Call("Push", oParagraph);
            // list
            createNumbering(oApi, data.achievements, "numbered", 22);

            // PLANS section
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Plans for the next year", 32, true);
            oDocument.Call("Push", oParagraph);
            // projects
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Projects:", 24);
            oDocument.Call("Push", oParagraph);
            // table
            var projects = data.plans.projects;
            oTable = createTable(oApi, projects.Count + 1, 2);
            fillTableHeaders(oTable, new string[] { "Name", "Deadline" }, 22);
            fillTableBody(oTable, projects, new string[] { "name", "deadline" }, 22);
            oDocument.Call("Push", oTable);
            // financial goals
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Financial goals:", 24);
            oDocument.Call("Push", oParagraph);
            // table
            var goals = data.plans.financial_goals;
            oTable = createTable(oApi, goals.Count + 1, 2);
            fillTableHeaders(oTable, new string[] { "Goal", "Value" }, 22);
            fillTableBody(oTable, goals, new string[] { "goal", "value" }, 22);
            oDocument.Call("Push", oTable);
            // marketing initiatives
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Marketing initiatives:", 24);
            oDocument.Call("Push", oParagraph);
            // list
            createNumbering(oApi, data.plans.marketing_initiatives, "bullet", 22);

            // save and close
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void addTextToParagraph(CValue oParagraph, string text, int fontSize, bool isBold = false, string jc = "left")
        {
            oParagraph.Call("AddText", text);
            oParagraph.Call("SetFontSize", fontSize);
            oParagraph.Call("SetBold", isBold);
            oParagraph.Call("SetJc", jc);
        }
        public static CValue createTable(CValue oApi, int rows, int cols, int borderColor = 200)
        {
            // create table
            CValue oTable = oApi.Call("CreateTable", cols, rows);
            // set table properties;
            oTable.Call("SetWidth", "percent", 100);
            oTable.Call("SetTableCellMarginTop", 200);
            oTable.Call("GetRow", 0).Call("SetBackgroundColor", 245, 245, 245);
            // set table borders;
            oTable.Call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
            return oTable;
        }
        public static CValue getTableCellParagraph(CValue oTable, int row, int col)
        {
            return oTable.Call("GetCell", row, col).Call("GetContent").Call("GetElement", 0);
        }

        public static void fillTableHeaders(CValue oTable, string[] data, int fontSize)
        {
            for (int i = 0; i < data.Length; i++)
            {
                CValue oParagraph = getTableCellParagraph(oTable, 0, i);
                addTextToParagraph(oParagraph, data[i], fontSize, true);
            }
        }

        public static void fillTableBody<T>(CValue oTable, List<T> data, string[] keys, int fontSize, int startRow = 1)
        {
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < keys.Length; col++)
                {
                    CValue oParagraph = getTableCellParagraph(oTable, row + startRow, col);
                    addTextToParagraph(oParagraph, (string)data[row].GetType().GetProperty(keys[col]).GetValue(data[row]), fontSize);
                }
            }
        }

        public static CValue createNumbering(CValue oApi, List<string> data, string numberingType, int fontSize)
        {
            CValue oDocument = oApi.Call("GetDocument");
            CValue oNumbering = oDocument.Call("CreateNumbering", numberingType);
            CValue oNumberingLevel = oNumbering.Call("GetLevel", 0);

            CValue oParagraph = CValue.CreateUndefined();
            foreach (string entry in data)
            {
                oParagraph = oApi.Call("CreateParagraph");
                oParagraph.Call("SetNumbering", oNumberingLevel);
                addTextToParagraph(oParagraph, entry, fontSize);
                oDocument.Call("Push", oParagraph);
            }
            // return the last oParagraph in numbering
            return oParagraph;
        }
    }

    // Define classes to represent the JSON structure
    public class YearData
    {
        public int year { get; set; }
        public FinancialData financials { get; set; }
        public List<string> achievements { get; set; }
        public PlansData plans { get; set; }
    }

    public class FinancialData
    {
        public int total_revenue { get; set; }
        public int total_expenses { get; set; }
        public int net_profit { get; set; }
        public List<QuarterlyData> quarterly_data { get; set; }
        public int r_d_expenses { get; set; }
        public int marketing_expenses { get; set; }
    }

    public class QuarterlyData
    {
        public string quarter { get; set; }
        public int revenue { get; set; }
        public int expenses { get; set; }
        public int net_profit { get; set; }
    }

    public class PlansData
    {
        public List<ProjectData> projects { get; set; }
        public List<FinancialGoalData> financial_goals { get; set; }
        public List<string> marketing_initiatives { get; set; }
    }

    public class ProjectData
    {
        public string name { get; set; }
        public string deadline { get; set; }
    }

    public class FinancialGoalData
    {
        public string goal { get; set; }
        public string value { get; set; }
    }
}
