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
        public static void Main()
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
            string jsonPath = resourcesDir + "/data/financial_system_response.json";
            string json = File.ReadAllText(jsonPath);
            YearData data = JsonSerializer.Deserialize<YearData>(json);

            // init docbuilder and create new docx file
            var doctype = (int)OfficeFileTypes.Document.DOCX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue document = api.Call("GetDocument");

            // DOCUMENT HEADER
            CValue paragraph = document.Call("GetElement", 0);
            AddTextToParagraph(paragraph, $"Annual Report for {data.year}", 44, true, "center");

            // FINANCIAL section
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Financial performance", 32, true);
            document.Call("Push", paragraph);
            // quarterly data
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Quarterly data:", 24);
            document.Call("Push", paragraph);
            // chart
            paragraph = api.Call("CreateParagraph");
            string[] chartKeys = { "revenue", "expenses", "net_profit" };
            var quarterlyData = data.financials.quarterly_data;
            CValue[] chartData = new CValue[chartKeys.Length];
            for (int i = 0; i < chartKeys.Length; i++)
            {
                chartData[i] = context.CreateArray(quarterlyData.Count);
                for (int j = 0; j < quarterlyData.Count; j++)
                {
                    chartData[i][j] = (int)quarterlyData[j].GetType().GetProperty(chartKeys[i]).GetValue(quarterlyData[j]);
                }
            }
            CValue[] chartNames = { "Revenue", "Expenses", "Net Profit" };
            CValue[] horValues = { "Q1", "Q2", "Q3", "Q4" };
            CValue chart = api.Call("CreateChart", "lineNormal", chartData, chartNames, horValues);
            chart.Call("SetSize", 170 * 36000, 90 * 36000);
            paragraph.Call("AddDrawing", chart);
            document.Call("Push", paragraph);
            // expenses
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Expenses:", 24);
            document.Call("Push", paragraph);
            // pie chart
            paragraph = api.Call("CreateParagraph");
            int rdExpenses = data.financials.r_d_expenses;
            int marketingExpenses = data.financials.marketing_expenses;
            int totalExpenses = data.financials.total_expenses;
            chartData = new CValue[1];
            chartData[0] = new CValue[] { rdExpenses, marketingExpenses, totalExpenses - (rdExpenses + marketingExpenses) };
            chartNames = new CValue[] { "Research and Development", "Marketing", "Other" };
            chart = api.Call("CreateChart", "pie", chartData, context.CreateArray(0), chartNames);
            chart.Call("SetSize", 170 * 36000, 90 * 36000);
            paragraph.Call("AddDrawing", chart);
            document.Call("Push", paragraph);
            // year totals
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Year total numbers:", 24);
            document.Call("Push", paragraph);
            // table
            CValue table = CreateTable(api, 2, 3);
            FillTableHeaders(table, new string[] { "Total revenue", "Total expenses", "Total net profit" }, 22);
            paragraph = GetTableCellParagraph(table, 1, 0);
            AddTextToParagraph(paragraph, data.financials.total_revenue.ToString(), 22);
            paragraph = GetTableCellParagraph(table, 1, 1);
            AddTextToParagraph(paragraph, data.financials.total_expenses.ToString(), 22);
            paragraph = GetTableCellParagraph(table, 1, 2);
            AddTextToParagraph(paragraph, data.financials.net_profit.ToString(), 22);
            document.Call("Push", table);

            // ACHIEVEMENTS section
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Achievements this year", 32, true);
            document.Call("Push", paragraph);
            // list
            CreateNumbering(api, data.achievements, "numbered", 22);

            // PLANS section
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Plans for the next year", 32, true);
            document.Call("Push", paragraph);
            // projects
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Projects:", 24);
            document.Call("Push", paragraph);
            // table
            var projects = data.plans.projects;
            table = CreateTable(api, projects.Count + 1, 2);
            FillTableHeaders(table, new string[] { "Name", "Deadline" }, 22);
            FillTableBody(table, projects, new string[] { "name", "deadline" }, 22);
            document.Call("Push", table);
            // financial goals
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Financial goals:", 24);
            document.Call("Push", paragraph);
            // table
            var goals = data.plans.financial_goals;
            table = CreateTable(api, goals.Count + 1, 2);
            FillTableHeaders(table, new string[] { "Goal", "Value" }, 22);
            FillTableBody(table, goals, new string[] { "goal", "value" }, 22);
            document.Call("Push", table);
            // marketing initiatives
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Marketing initiatives:", 24);
            document.Call("Push", paragraph);
            // list
            CreateNumbering(api, data.plans.marketing_initiatives, "bullet", 22);

            // save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void AddTextToParagraph(CValue paragraph, string text, int fontSize, bool isBold = false, string jc = "left")
        {
            paragraph.Call("AddText", text);
            paragraph.Call("SetFontSize", fontSize);
            paragraph.Call("SetBold", isBold);
            paragraph.Call("SetJc", jc);
        }

        public static CValue CreateTable(CValue api, int rows, int cols, int borderColor = 200)
        {
            // create table
            CValue table = api.Call("CreateTable", cols, rows);
            // set table properties;
            table.Call("SetWidth", "percent", 100);
            table.Call("SetTableCellMarginTop", 200);
            table.Call("GetRow", 0).Call("SetBackgroundColor", 245, 245, 245);
            // set table borders;
            table.Call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
            return table;
        }

        public static CValue GetTableCellParagraph(CValue table, int row, int col)
        {
            return table.Call("GetCell", row, col).Call("GetContent").Call("GetElement", 0);
        }

        public static void FillTableHeaders(CValue table, string[] data, int fontSize)
        {
            for (int i = 0; i < data.Length; i++)
            {
                CValue paragraph = GetTableCellParagraph(table, 0, i);
                AddTextToParagraph(paragraph, data[i], fontSize, true);
            }
        }

        public static void FillTableBody<T>(CValue table, List<T> data, string[] keys, int fontSize, int startRow = 1)
        {
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < keys.Length; col++)
                {
                    CValue paragraph = GetTableCellParagraph(table, row + startRow, col);
                    AddTextToParagraph(paragraph, (string)data[row].GetType().GetProperty(keys[col]).GetValue(data[row]), fontSize);
                }
            }
        }

        public static CValue CreateNumbering(CValue api, List<string> data, string numberingType, int fontSize)
        {
            CValue document = api.Call("GetDocument");
            CValue numbering = document.Call("CreateNumbering", numberingType);
            CValue numberingLevel = numbering.Call("GetLevel", 0);

            CValue paragraph = CValue.CreateUndefined();
            foreach (string entry in data)
            {
                paragraph = api.Call("CreateParagraph");
                paragraph.Call("SetNumbering", numberingLevel);
                AddTextToParagraph(paragraph, entry, fontSize);
                document.Call("Push", paragraph);
            }
            // return the last paragraph in numbering
            return paragraph;
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
