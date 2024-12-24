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

        createDevelopmentPlan(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createDevelopmentPlan(String resultPath, String resourcesDir) throws Exception {
        // parse JSON
        String jsonPath = resourcesDir + "/data/hrms_response.json";
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

        // TITLE PAGE
        // header
        CDocBuilderValue paragraph = document.call("GetElement", 0);
        addTextToParagraph(paragraph, "Employee Development Plan for 2024", 48, true, "center");
        paragraph.call("SetSpacingBefore", 5000);
        paragraph.call("SetSpacingAfter", 500);
        // employee name
        paragraph = api.call("CreateParagraph");
        JSONObject employee = (JSONObject)data.get("employee");
        addTextToParagraph(paragraph, employee.get("name").toString(), 36, false, "center");
        document.call("Push", paragraph);
        // employee position and department
        paragraph = api.call("CreateParagraph");
        String employeeInfo = String.format("Position: %s\nDepartment: %s", employee.get("position").toString(), employee.get("department").toString());
        addTextToParagraph(paragraph, employeeInfo, 24, false, "center");
        paragraph.call("AddPageBreak");
        document.call("Push", paragraph);

        // COMPETENCIES SECION
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Competencies", 32, true);
        document.call("Push", paragraph);
        // technical skills sub-header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Technical skills:", 24);
        document.call("Push", paragraph);
        // technical skills table
        JSONObject competencies = (JSONObject)data.get("competencies");
        JSONArray technicalSkills = (JSONArray)competencies.get("technical_skills");
        CDocBuilderValue table = createTable(api, technicalSkills.size() + 1, 2);
        fillTableHeaders(table, new String[] { "Skill", "Level" }, 22);
        fillTableBody(table, technicalSkills, new String[] { "name", "level" }, 22);
        document.call("Push", table);
        // soft skills sub-header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Soft skills:", 24);
        document.call("Push", paragraph);
        // soft skills table
        JSONArray softSkills = (JSONArray)competencies.get("soft_skills");
        table = createTable(api, softSkills.size() + 1, 2);
        fillTableHeaders(table, new String[] { "Skill", "Level" }, 22);
        fillTableBody(table, softSkills, new String[] { "name", "level" }, 22);
        document.call("Push", table);

        // DEVELOPMENT AREAS section
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Development areas", 32, true);
        document.call("Push", paragraph);
        // list
        createNumbering(api, (JSONArray)data.get("development_areas"), "numbered", 22);

        // GOALS section
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Goals for next year", 32, true);
        document.call("Push", paragraph);
        // numbering
        paragraph = createNumbering(api, (JSONArray)data.get("goals_next_year"), "numbered", 22);
        // add a page break after the last paragraph
        paragraph.call("AddPageBreak");

        // RESOURCES section
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Recommended resources", 32, true);
        document.call("Push", paragraph);
        // table
        JSONArray resources = (JSONArray)data.get("resources");
        table = createTable(api, resources.size() + 1, 3);
        fillTableHeaders(table, new String[] { "Name", "Provider", "Duration" }, 22);
        fillTableBody(table, resources, new String[] { "name", "provider", "duration" }, 22);
        document.call("Push", table);

        // FEEDBACK section
        // header
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Feedback", 32, true);
        document.call("Push", paragraph);
        // manager's feedback
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Manager's feedback:", 24, false);
        document.call("Push", paragraph);
        paragraph = api.call("CreateParagraph");
        // make blank lines string
        StringBuilder blankLinesBuilder = new StringBuilder(280);
        for (int i = 0; i < blankLinesBuilder.capacity(); i++) {
            blankLinesBuilder.append('_');
        }
        String blankLines = blankLinesBuilder.toString();
        addTextToParagraph(paragraph, blankLines, 24, false);
        document.call("Push", paragraph);
        // employees's feedback
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, "Employee's feedback:", 24, false);
        document.call("Push", paragraph);
        paragraph = api.call("CreateParagraph");
        addTextToParagraph(paragraph, blankLines, 24, false);
        document.call("Push", paragraph);

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
