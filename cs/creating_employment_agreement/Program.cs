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
    public class CreatingEmploymentAgreement
    {
        const int defaultFontSize = 24;
        const string defaultJc = "both";
        static readonly string[] signerData = {
            "Name: __________________________",
            "Signature: _______________________",
            "Date: ___________________________"
        };

        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.pdf";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateEmploymentAgreement(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateEmploymentAgreement(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string jsonPath = resourcesDir + "/data/employment_agreement_data.json";
            string json = File.ReadAllText(jsonPath);
            JsonData data = JsonSerializer.Deserialize<JsonData>(json);

            // init docbuilder and create new docx file
            var doctype = (int)OfficeFileTypes.Document.OFORM_PDF;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue document = api.Call("GetDocument");

            // DOCUMENT STYLE
            CValue paraPr = document.Call("GetDefaultParaPr");
            paraPr.Call("SetJc", "both");
            CValue textPr = document.Call("GetDefaultTextPr");
            textPr.Call("SetFontSize", 24);
            textPr.Call("SetFontFamily", "Times New Roman");

            // DOCUMENT HEADER
            CValue header = document.Call("GetElement", 0);
            header.Call("AddText", "EMPLOYMENT AGREEMENT");
            header.Call("SetFontSize", 28);
            header.Call("SetBold", true);

            CValue headerDesc = CreateParagraph(
                api,
                $"This Employment Agreement (\"Agreement\") is made and entered into on {data.date} by and between:"
            );
            SetSpacingAfter(headerDesc, 50);
            document.Call("Push", headerDesc);

            // PARTICIPANTS OF THE DOCUMENT
            CValue participants = CreateParagraph(api, "", false, defaultFontSize, "left");
            AddParticipantToParagraph(
                api,
                participants,
                "Employer",
                $"{data.employer.name}, located at {data.employer.address}."
            );
            participants.Call("AddLineBreak");
            AddParticipantToParagraph(
                api,
                participants,
                "Employee",
                $"{data.employee.full_name}, residing at {data.employee.address}."
            );
            document.Call("Push", participants);
            document.Call("Push", CreateParagraph(api, "The parties agree to the following terms and conditions:"));

            // AGREEMENT CONDITIONS
            // Create numbering
            CValue numbering = document.Call("CreateNumbering", "numbered");
            CValue numberingLvl = numbering.Call("GetLevel", 0);
            numberingLvl.Call("SetCustomType", "decimal", "%1.", "left");
            numberingLvl.Call("SetSuff", "space");

            // Position and duties
            document.Call("Push", CreateNumberedSection(api, "POSITION AND DUTIES", numberingLvl));
            document.Call(
                "Push",
                CreateConditionsDescParagraph(
                    api,
                    $"The Employee is hired as {data.position_and_duties.job_title}. " +
                    "The Employee shall perform their duties as outlined by the Employer and comply with all applicable policies and guidelines."
                )
            );

            // Compensation
            document.Call("Push", CreateNumberedSection(api, "COMPENSATION", numberingLvl));
            document.Call(
                "Push",
                CreateConditionsDescParagraph(
                    api,
                    $"The Employee will receive a salary of {data.compensation.salary.ToString()} " +
                    $"{data.compensation.currency} {data.compensation.frequency} ({data.compensation.type}), " +
                    "payable in accordance with the Employer's payroll schedule and subject to lawful deductions."
                )
            );

            // Probationary period
            document.Call("Push", CreateNumberedSection(api, "PROBATIONARY PERIOD", numberingLvl));
            document.Call(
                "Push",
                CreateConditionsDescParagraph(
                    api,
                    $"The Employee will serve a probationary period of {data.probationary_period.duration}. " +
                    "During this period, the Employer may terminate this Agreement with " +
                    $"{data.probationary_period.terminate} days' notice if performance is deemed unsatisfactory."
                )
            );

            // Work conditions
            document.Call("Push", CreateNumberedSection(api, "WORK CONDITIONS", numberingLvl));
            CValue conditionsText = CreateConditionsDescParagraph(
                api,
                "The following terms apply to the Employee's working conditions:"
            );
            SetSpacingAfter(conditionsText, 50);
            document.Call("Push", conditionsText);

            // Create bullet numbering
            CValue bulletNumbering = document.Call("CreateNumbering", "bullet");
            CValue bulletNumLvl = bulletNumbering.Call("GetLevel", 0);

            document.Call(
                "Push",
                CreateWorkCondition(api, "Working Hours", data.work_conditions.working_hours, bulletNumLvl, true)
            );
            document.Call(
                "Push",
                CreateWorkCondition(api, "Work Schedule", data.work_conditions.work_schedule, bulletNumLvl, true)
            );
            document.Call(
                "Push",
                CreateWorkCondition(api, "Benefits", string.Join(", ", data.work_conditions.benefits), bulletNumLvl, true)
            );
            document.Call(
                "Push",
                CreateWorkCondition(api, "Other terms", string.Join(", ", data.work_conditions.other_terms), bulletNumLvl, false)
            );

            // TERMINATION
            document.Call("Push", CreateNumberedSection(api, "TERMINATION", numberingLvl));
            document.Call(
                "Push",
                CreateConditionsDescParagraph(
                    api,
                    $"Either party may terminate this Agreement by providing {data.termination.notice_period} written notice. " +
                    "The Employer reserves the right to terminate employment immediately for cause, including but not limited to misconduct or breach of Agreement."
                )
            );

            // GOVERNING LAW
            document.Call("Push", CreateNumberedSection(api, "GOVERNING LAW", numberingLvl));
            document.Call(
                "Push",
                CreateConditionsDescParagraph(
                    api,
                    $"This Agreement is governed by the laws of {data.governing_law.jurisdiction}, " +
                    "and any disputes arising under this Agreement will be resolved in accordance with these laws."
                )
            );

            // ENTIRE AGREEMENT
            document.Call("Push", CreateNumberedSection(api, "ENTIRE AGREEMENT", numberingLvl));
            document.Call(
                "Push",
                CreateConditionsDescParagraph(
                    api,
                    "This document constitutes the entire Agreement between the parties and supersedes all prior agreements. " +
                    "Any amendments must be made in writing and signed by both parties."
                )
            );

            // Signatures
            CValue table = api.Call("CreateTable", 2, 2);
            // set table properties
            table.Call("SetWidth", "percent", 100);
            // fill table
            CValue tableTitle = table.Call("GetRow", 0);
            CValue titleParagraph = tableTitle.Call("MergeCells").Call("GetContent").Call("GetElement", 0);
            titleParagraph.Call("Push", CreateRun(api, "SIGNATURES", true, 24));
            FillSigner(api, table.Call("GetCell", 1, 0), "Employer");
            FillSigner(api, table.Call("GetCell", 1, 1), "Employee");
            document.Call("Push", table);

            // save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static CValue CreateParagraph(CValue api, string text, bool isBold = false, int fontSize = defaultFontSize, string jc = defaultJc)
        {
            CValue paragraph = api.Call("CreateParagraph");
            paragraph.Call("AddText", text);
            paragraph.Call("SetBold", isBold);
            if (fontSize != defaultFontSize) {
                paragraph.Call("SetFontSize", fontSize);
            }
            if (jc != defaultJc ) {
                paragraph.Call("SetJc", jc);
            }
            return paragraph;
        }

        public static CValue CreateRun(CValue api, string text, bool isBold = false, int fontSize = defaultFontSize)
        {
            CValue run = api.Call("CreateRun");
            run.Call("AddText", text);
            run.Call("SetBold", isBold);
            if (fontSize != defaultFontSize) {
                run.Call("SetFontSize", fontSize);
            }
            return run;
        }

        public static void SetNumbering(CValue paragraph, CValue numLvl)
        {
            paragraph.Call("SetNumbering", numLvl);
        }

        public static void SetSpacingAfter(CValue paragraph, int spacing)
        {
            paragraph.Call("SetSpacingAfter", spacing);
        }

        public static CValue CreateConditionsDescParagraph(CValue api, string text)
        {
            // create paragraph with first line indentation
            CValue paragraph = CreateParagraph(api, text);
            paragraph.Call("SetIndFirstLine", 400);
            return paragraph;
        }

        public static void AddParticipantToParagraph(CValue api, CValue paragraph, string pType, string details)
        {
            paragraph.Call("Push", CreateRun(api, pType + ": ", true));
            paragraph.Call("Push", CreateRun(api, details));
        }

        public static CValue CreateNumberedSection(CValue api, string text, CValue numLvl)
        {
            CValue paragraph = CreateParagraph(api, text, true);
            SetNumbering(paragraph, numLvl);
            SetSpacingAfter(paragraph, 50);
            return paragraph;
        }

        public static CValue CreateWorkCondition(CValue api, string title, string text, CValue numLvl, bool setSpacing = false)
        {
            CValue paragraph = api.Call("CreateParagraph");
            SetNumbering(paragraph, numLvl);
            if (setSpacing) {
                SetSpacingAfter(paragraph, 20);
            }
            paragraph.Call("SetJc", "left");
            paragraph.Call("Push", CreateRun(api, title + ": ", true));
            paragraph.Call("Push", CreateRun(api, text));
            return paragraph;
        }

        public static void FillSigner(CValue api, CValue cell, string title)
        {
            CValue paragraph = cell.Call("GetContent").Call("GetElement", 0);
            paragraph.Call("SetJc", "left");
            paragraph.Call("Push", CreateRun(api, title, true));

            foreach (string text in signerData) {
                paragraph.Call("AddLineBreak");
                paragraph.Call("Push", CreateRun(api, text));
            }
        }
    }

    // Define classes to represent the JSON structure
    public class JsonData
    {
       public string date { get; set; }
       public Employer employer { get; set; }
       public Employee employee { get; set; }
       public PositionAndDuties position_and_duties { get; set; }
       public Compensation compensation { get; set; }
       public ProbationaryPeriod probationary_period { get; set; }
       public WorkConditions work_conditions { get; set; }
       public Termination termination { get; set; }
       public GoverningLaw governing_law { get; set; }
    }

    public class Employer
    {
        public string name { get; set; }
        public string address { get; set; }
    }

    public class Employee
    {
        public string full_name { get; set; }
        public string address { get; set; }
    }

    public class PositionAndDuties
    {
        public string job_title { get; set; }
    }

    public class Compensation
    {
        public int salary { get; set; }
        public string currency { get; set; }
        public string frequency { get; set; }
        public string type { get; set; }
    }

    public class ProbationaryPeriod
    {
        public string duration { get; set; }
        public string terminate { get; set; }
    }

    public class WorkConditions
    {
        public string working_hours { get; set; }
        public string work_schedule { get; set; }
        public List<string> benefits { get; set; }
        public List<string> other_terms { get; set; }
    }

    public class Termination
    {
        public string notice_period { get; set; }
    }

    public class GoverningLaw
    {
        public string jurisdiction { get; set; }
    }
}
