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

            CreateDevelopmentPlan(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateDevelopmentPlan(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string jsonPath = resourcesDir + "/data/hrms_response.json";
            string json = File.ReadAllText(jsonPath);
            EmployeeData data = JsonSerializer.Deserialize<EmployeeData>(json);

            // init docbuilder and create new docx file
            var doctype = (int)OfficeFileTypes.Document.DOCX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue document = api.Call("GetDocument");

            // TITLE PAGE
            // header
            CValue paragraph = document.Call("GetElement", 0);
            AddTextToParagraph(paragraph, "Employee Development Plan for 2024", 48, true, "center");
            paragraph.Call("SetSpacingBefore", 5000);
            paragraph.Call("SetSpacingAfter", 500);
            // employee name
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, data.employee.name, 36, false, "center");
            document.Call("Push", paragraph);
            // employee position and department
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, $"Position: {data.employee.position}\nDepartment: {data.employee.department}", 24, false, "center");
            paragraph.Call("AddPageBreak");
            document.Call("Push", paragraph);

            // COMPETENCIES SECION
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Competencies", 32, true);
            document.Call("Push", paragraph);
            // technical skills sub-header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Technical skills:", 24);
            document.Call("Push", paragraph);
            // technical skills table
            var technicalSkills = data.competencies.technical_skills;
            CValue table = CreateTable(api, technicalSkills.Count + 1, 2);
            FillTableHeaders(table, new string[] { "Skill", "Level" }, 22);
            FillTableBody(table, technicalSkills, new string[] { "name", "level" }, 22);
            document.Call("Push", table);
            // soft skills sub-header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Soft skills:", 24);
            document.Call("Push", paragraph);
            // soft skills table
            var softSkills = data.competencies.soft_skills;
            table = CreateTable(api, softSkills.Count + 1, 2);
            FillTableHeaders(table, new string[] { "Skill", "Level" }, 22);
            FillTableBody(table, softSkills, new string[] { "name", "level" }, 22);
            document.Call("Push", table);

            // DEVELOPMENT AREAS section
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Development areas", 32, true);
            document.Call("Push", paragraph);
            // list
            CreateNumbering(api, data.development_areas, "numbered", 22);

            // GOALS section
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Goals for next year", 32, true);
            document.Call("Push", paragraph);
            // numbering
            paragraph = CreateNumbering(api, data.goals_next_year, "numbered", 22);
            // add a page break after the last paragraph
            paragraph.Call("AddPageBreak");

            // RESOURCES section
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Recommended resources", 32, true);
            document.Call("Push", paragraph);
            // table
            var resources = data.resources;
            table = CreateTable(api, resources.Count + 1, 3);
            FillTableHeaders(table, new string[] { "Name", "Provider", "Duration" }, 22);
            FillTableBody(table, resources, new string[] { "name", "provider", "duration" }, 22);
            document.Call("Push", table);

            // FEEDBACK section
            // header
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Feedback", 32, true);
            document.Call("Push", paragraph);
            // manager's feedback
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Manager's feedback:", 24, false);
            document.Call("Push", paragraph);
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, new string('_', 280), 24, false);
            document.Call("Push", paragraph);
            // employees's feedback
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Employee's feedback:", 24, false);
            document.Call("Push", paragraph);
            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, new string('_', 280), 24, false);
            document.Call("Push", paragraph);

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
