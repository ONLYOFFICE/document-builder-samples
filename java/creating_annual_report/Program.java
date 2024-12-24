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

import docbuilder.*;

import java.io.FileReader;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class Program {
    public static void main(String[] args) throws Exception {
        String resultPath = "result.docx";
        String resourcesDir = "../../resources";

        createAnnualReport(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createAnnualReport(String resultPath, String resourcesDir) throws Exception {
        // parse JSON
        String jsonPath = resourcesDir + "/data/financial_system_response.json";
        JSONObject data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));

        // init docbuilder and create new docx file
        int doctype = FileTypes.Document.DOCX;
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");
        CDocBuilderValue document = api.call("GetDocument");

        // DOCUMENT HEADER
        CDocBuilderValue paragraph = document.call("GetElement", 0);
        addTextToParagraph(paragraph, "Annual Report for " + data.get("year").toString(), 44, true, "center");

        // FINANCIAL section
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Financial performance", 32, true);
        document.call("Push", paragraph);
        // quarterly data
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Quarterly data:", 24);
        document.call("Push", paragraph);
        // chart
        paragraph = api.call("CreateParagraph");
        String[] chartKeys = { "revenue", "expenses", "net_profit" };
        JSONObject financials = (JSONObject)data.get("financials");
        JSONArray quarterlyData = (JSONArray)financials.get("quarterly_data");
        CDocBuilderValue[] chartData = new CDocBuilderValue[chartKeys.length];
        for (int i = 0; i < chartKeys.length; i++) {
            chartData[i] = context.createArray(quarterlyData.size());
            for (int j = 0; j < quarterlyData.size(); j++) {
                chartData[i].set(j, (int)(long)((JSONObject)quarterlyData.get(j)).get(chartKeys[i]));
            }
        }
        CDocBuilderValue chartNames = new CDocBuilderValue(new Object[] { "Revenue", "Expenses", "Net Profit" });
        CDocBuilderValue horValues = new CDocBuilderValue(new Object[] { "Q1", "Q2", "Q3", "Q4" });
        CDocBuilderValue chart = api.call("CreateChart", "lineNormal", chartData, chartNames, horValues);
        chart.call("SetSize", 170 * 36000, 90 * 36000);
        paragraph.call("AddDrawing", chart);
        document.call("Push", paragraph);
        // expenses
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Expenses:", 24);
        document.call("Push", paragraph);
        // pie chart
        paragraph = api.call("CreateParagraph");
        int rdExpenses = (int)(long)financials.get("r_d_expenses");
        int marketingExpenses = (int)(long)financials.get("marketing_expenses");
        int totalExpenses = (int)(long)financials.get("total_expenses");
        chartData = new CDocBuilderValue[1];
        chartData[0] = new CDocBuilderValue(new Object[] { rdExpenses, marketingExpenses, totalExpenses - (rdExpenses + marketingExpenses) });
        chartNames = new CDocBuilderValue(new Object[] { "Research and Development", "Marketing", "Other" });;
        chart = api.call("CreateChart", "pie", chartData, context.createArray(0), chartNames);
        chart.call("SetSize", 170 * 36000, 90 * 36000);
        paragraph.call("AddDrawing", chart);
        document.call("Push", paragraph);
        // year totals
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Year total numbers:", 24);
        document.call("Push", paragraph);
        // table
        CDocBuilderValue table = createTable(api, 2, 3);
        fillTableHeaders(table, new String[] { "Total revenue", "Total expenses", "Total net profit" }, 22);
        paragraph = getTableCellParagraph(table, 1, 0);
        addTextToParagraph(paragraph, financials.get("total_revenue").toString(), 22);
        paragraph = getTableCellParagraph(table, 1, 1);
        addTextToParagraph(paragraph, financials.get("total_expenses").toString(), 22);
        paragraph = getTableCellParagraph(table, 1, 2);
        addTextToParagraph(paragraph, financials.get("net_profit").toString(), 22);
        document.call("Push", table);

        // ACHIEVEMENTS section
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Achievements this year", 32, true);
        document.call("Push", paragraph);
        // list
        createNumbering(api, (JSONArray)data.get("achievements"), "numbered", 22);

        // PLANS section
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Plans for the next year", 32, true);
        document.call("Push", paragraph);
        // projects
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Projects:", 24);
        document.call("Push", paragraph);
        // table
        JSONObject plans = (JSONObject)data.get("plans");
        JSONArray projects = (JSONArray)plans.get("projects");
        table = createTable(api, projects.size() + 1, 2);
        fillTableHeaders(table, new String[] { "Name", "Deadline" }, 22);
        fillTableBody(table, projects, new String[] { "name", "deadline" }, 22);
        document.call("Push", table);
        // financial goals
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Financial goals:", 24);
        document.call("Push", paragraph);
        // table
        JSONArray goals = (JSONArray)plans.get("financial_goals");
        table = createTable(api, goals.size() + 1, 2);
        fillTableHeaders(table, new String[] { "Goal", "Value" }, 22);
        fillTableBody(table, goals, new String[] { "goal", "value" }, 22);
        document.call("Push", table);
        // marketing initiatives
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Marketing initiatives:", 24);
        document.call("Push", paragraph);
        // list
        createNumbering(api, (JSONArray)plans.get("marketing_initiatives"), "bullet", 22);

        // save and close
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }

    public static void addTextToParagraph(CDocBuilderValue paragraph, String text, int fontSize, boolean isBold, String jc) {
        paragraph.call("AddText", text);
        paragraph.call("SetFontSize", fontSize);
        paragraph.call("SetBold", isBold);
        paragraph.call("SetJc", jc);
    }

    public static void addTextToParagraph(CDocBuilderValue paragraph, String text, int fontSize, boolean isBold) {
        addTextToParagraph(paragraph, text, fontSize, isBold, "left");
    }

    public static void addTextToParagraph(CDocBuilderValue paragraph, String text, int fontSize) {
        addTextToParagraph(paragraph, text, fontSize, false, "left");
    }

    public static CDocBuilderValue createTable(CDocBuilderValue api, int rows, int cols, int borderColor) {
        // create table
        CDocBuilderValue table = api.call("CreateTable", cols, rows);
        // set table properties;
        table.call("SetWidth", "percent", 100);
        table.call("SetTableCellMarginTop", 200);
        table.call("GetRow", 0).call("SetBackgroundColor", 245, 245, 245);
        // set table borders;
        table.call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
        table.call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
        table.call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
        table.call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
        table.call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
        table.call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
        return table;
    }

    public static CDocBuilderValue createTable(CDocBuilderValue api, int rows, int cols) {
        return createTable(api, rows, cols, 200);
    }

    public static CDocBuilderValue getTableCellParagraph(CDocBuilderValue table, int row, int col) {
        return table.call("GetCell", row, col).call("GetContent").call("GetElement", 0);
    }

    public static void fillTableHeaders(CDocBuilderValue table, String[] data, int fontSize) {
        for (int i = 0; i < data.length; i++) {
            CDocBuilderValue paragraph = getTableCellParagraph(table, 0, i);
            addTextToParagraph(paragraph, data[i], fontSize, true);
        }
    }

    public static void fillTableBody(CDocBuilderValue table, JSONArray data, String[] keys, int fontSize, int startRow) {
        for (int row = 0; row < data.size(); row++) {
            for (int col = 0; col < keys.length; col++) {
                CDocBuilderValue paragraph = getTableCellParagraph(table, row + startRow, col);
                addTextToParagraph(paragraph, ((JSONObject)data.get(row)).get(keys[col]).toString(), fontSize);
            }
        }
    }

    public static void fillTableBody(CDocBuilderValue table, JSONArray data, String[] keys, int fontSize) {
        fillTableBody(table, data, keys, fontSize, 1);
    }

    public static CDocBuilderValue createNumbering(CDocBuilderValue api, JSONArray data, String numberingType, int fontSize) {
        CDocBuilderValue document = api.call("GetDocument");
        CDocBuilderValue numbering = document.call("CreateNumbering", numberingType);
        CDocBuilderValue numberingLevel = numbering.call("GetLevel", 0);

        CDocBuilderValue paragraph = CDocBuilderValue.createUndefined();
        for (Object entry : data) {
            paragraph = api.call("CreateParagraph");
            paragraph.call("SetNumbering", numberingLevel);
            addTextToParagraph(paragraph, entry.toString(), fontSize);
            document.call("Push", paragraph);
        }
        // return the last paragraph in numbering
        return paragraph;
    }
}
