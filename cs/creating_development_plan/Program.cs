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
            string json_path = resourcesDir + "/data/hrms_response.json";
            string json = File.ReadAllText(json_path);
            EmployeeData data = JsonSerializer.Deserialize<EmployeeData>(json);

            // init docbuilder and create new docx file
            var doctype = (int)OfficeFileTypes.Document.DOCX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.CreateFile(doctype);

            CContext oContext = oBuilder.GetContext();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oDocument = oApi.Call("GetDocument");

            // TITLE PAGE
            // header
            CValue oParagraph = oDocument.Call("GetElement", 0);
            addTextToParagraph(oParagraph, "Employee Development Plan for 2024", 48, true, "center");
            oParagraph.Call("SetSpacingBefore", 5000);
            oParagraph.Call("SetSpacingAfter", 500);
            // employee name
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, data.employee.name, 36, false, "center");
            oDocument.Call("Push", oParagraph);
            // employee position and department
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, $"Position: {data.employee.position}\nDepartment: {data.employee.department}", 24, false, "center");
            oParagraph.Call("AddPageBreak");
            oDocument.Call("Push", oParagraph);

            // COMPETENCIES SECION
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Competencies", 32, true);
            oDocument.Call("Push", oParagraph);
            // technical skills sub-header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Technical skills:", 24);
            oDocument.Call("Push", oParagraph);
            // technical skills table
            var technicalSkills = data.competencies.technical_skills;
            CValue oTable = createTable(oApi, technicalSkills.Count + 1, 2);
            fillTableHeaders(oTable, new string[] { "Skill", "Level" }, 22);
            fillTableBody(oTable, technicalSkills, new string[] { "name", "level" }, 22);
            oDocument.Call("Push", oTable);
            // soft skills sub-header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Soft skills:", 24);
            oDocument.Call("Push", oParagraph);
            // soft skills table
            var softSkills = data.competencies.soft_skills;
            oTable = createTable(oApi, softSkills.Count + 1, 2);
            fillTableHeaders(oTable, new string[] { "Skill", "Level" }, 22);
            fillTableBody(oTable, softSkills, new string[] { "name", "level" }, 22);
            oDocument.Call("Push", oTable);

            // DEVELOPMENT AREAS section
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Development areas", 32, true);
            oDocument.Call("Push", oParagraph);
            // list
            createNumbering(oApi, data.development_areas, "numbered", 22);

            // GOALS section
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Goals for next year", 32, true);
            oDocument.Call("Push", oParagraph);
            // numbering
            oParagraph = createNumbering(oApi, data.goals_next_year, "numbered", 22);
            // add a page break after the last paragraph
            oParagraph.Call("AddPageBreak");

            // RESOURCES section
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Recommended resources", 32, true);
            oDocument.Call("Push", oParagraph);
            // table
            var resources = data.resources;
            oTable = createTable(oApi, resources.Count + 1, 3);
            fillTableHeaders(oTable, new string[] { "Name", "Provider", "Duration" }, 22);
            fillTableBody(oTable, resources, new string[] { "name", "provider", "duration" }, 22);
            oDocument.Call("Push", oTable);

            // FEEDBACK section
            // header
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Feedback", 32, true);
            oDocument.Call("Push", oParagraph);
            // manager"s feedback
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Manager's feedback:", 24, false);
            oDocument.Call("Push", oParagraph);
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, new string('_', 280), 24, false);
            oDocument.Call("Push", oParagraph);
            // employees"s feedback
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Employee's feedback:", 24, false);
            oDocument.Call("Push", oParagraph);
            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, new string('_', 280), 24, false);
            oDocument.Call("Push", oParagraph);

            // Save file and close DocBuilder
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
    public class EmployeeData
    {
        public EmployeePersonalData employee { get; set; }
        public CompetenciesData competencies { get; set; }
        public List<string> development_areas { get; set; }
        public List<string> goals_next_year { get; set; }
        public List<ResourceData> resources { get; set; }
    }

    public class EmployeePersonalData
    {
        public string name { get; set; }
        public string position { get; set; }
        public string department { get; set; }
    }

    public class CompetenciesData
    {
        public List<SkillData> technical_skills { get; set; }
        public List<SkillData> soft_skills { get; set; }
    }

    public class SkillData
    {
        public string name { get; set; }
        public string level { get; set; }
    }

    public class ResourceData
    {
        public string name { get; set; }
        public string provider { get; set; }
        public string duration { get; set; }
    }
}
