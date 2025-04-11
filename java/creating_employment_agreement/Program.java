/**
 *
 * (c) Copyright Ascensio System SIA 2025
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
    static int defaultFontSize = 24;
    static String defaultJc = "both";
    static String[] signerData = {
        "Name: __________________________",
        "Signature: _______________________",
        "Date: ___________________________"
    };

    public static void main(String[] args) throws Exception {
        String resultPath = "result.pdf";
        String resourcesDir = "../../resources";

        createEmploymentAgreement(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createEmploymentAgreement(String resultPath, String resourcesDir) throws Exception {
        // parse JSON
        String jsonPath = resourcesDir + "/data/employment_agreement_data.json";
        JSONObject data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));

        // init docbuilder and create new docx file
        int doctype = FileTypes.Document.OFORM_PDF;
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");
        CDocBuilderValue document = api.call("GetDocument");

        // DOCUMENT STYLE
        CDocBuilderValue paraPr = document.call("GetDefaultParaPr");
        paraPr.call("SetJc", defaultJc);
        CDocBuilderValue textPr = document.call("GetDefaultTextPr");
        textPr.call("SetFontSize", defaultFontSize);
        textPr.call("SetFontFamily", "Times New Roman");

        // DOCUMENT HEADER
        CDocBuilderValue header = document.call("GetElement", 0);
        header.call("AddText", "EMPLOYMENT AGREEMENT");
        header.call("SetFontSize", 28);
        header.call("SetBold", true);

        CDocBuilderValue headerDesc = createParagraph(
            api,
            String.format(
                "This Employment Agreement (\"Agreement\") is made and entered into on %s by and between:",
                data.get("date").toString()
            ),
            false
        );
        setSpacingAfter(headerDesc, 50);
        document.call("Push", headerDesc);

        // PARTICIPANTS OF THE DOCUMENT
        CDocBuilderValue participants = createParagraph(api, "", false, defaultFontSize, "left");
        JSONObject employer = (JSONObject)data.get("employer");
        addParticipantToParagraph(
            api,
            participants,
            "Employer",
            String.format("%s, located at %s.", employer.get("name").toString(), employer.get("address").toString())
        );
        participants.call("AddLineBreak");
        JSONObject employee = (JSONObject)data.get("employee");
        addParticipantToParagraph(
            api,
            participants,
            "Employee",
            String.format("%s, residing at %s.", employee.get("full_name").toString(), employee.get("address").toString())
        );
        document.call("Push", participants);
        document.call("Push", createParagraph(api, "The parties agree to the following terms and conditions:", false));

        // AGREEMENT CONDITIONS
        // Create numbering
        CDocBuilderValue numbering = document.call("CreateNumbering", "numbered");
        CDocBuilderValue numberingLvl = numbering.call("GetLevel", 0);
        numberingLvl.call("SetCustomType", "decimal", "%1.", "left");
        numberingLvl.call("SetSuff", "space");

        // Position and duties
        document.call("Push", createNumberedSection(api, "POSITION AND DUTIES", numberingLvl));
        document.call(
            "Push",
            createConditionsDescParagraph(
                api,
                String.format(
                    "The Employee is hired as %s. The Employee shall perform their duties as outlined by the Employer and comply with all applicable policies and guidelines.",
                    ((JSONObject)data.get("position_and_duties")).get("job_title")
                )
            )
        );

        // Compensation
        document.call("Push", createNumberedSection(api, "COMPENSATION", numberingLvl));
        JSONObject compensation = (JSONObject)data.get("compensation");
        document.call(
            "Push",
            createConditionsDescParagraph(
                api,
                String.format(
                    "The Employee will receive a salary of %s %s %s (%s), payable in accordance with the Employer's payroll schedule and subject to lawful deductions.",
                    compensation.get("salary").toString(), compensation.get("currency").toString(), compensation.get("frequency").toString(), compensation.get("type").toString()
                )
            )
        );

        // Probationary period
        document.call("Push", createNumberedSection(api, "PROBATIONARY PERIOD", numberingLvl));
        JSONObject probPeriod = (JSONObject)data.get("probationary_period");
        document.call(
            "Push",
            createConditionsDescParagraph(
                api,
                String.format(
                    "The Employee will serve a probationary period of %s. During this period, the Employer may terminate this Agreement with %s days' notice if performance is deemed unsatisfactory.",
                    probPeriod.get("duration").toString(), probPeriod.get("terminate").toString()
                )
            )
        );

        // Work conditions
        document.call("Push", createNumberedSection(api, "WORK CONDITIONS", numberingLvl));
        CDocBuilderValue conditionsText = createConditionsDescParagraph(
            api,
            "The following terms apply to the Employee's working conditions:"
        );
        setSpacingAfter(conditionsText, 50);
        document.call("Push", conditionsText);

        // Create bullet numbering
        CDocBuilderValue bulletNumbering = document.call("CreateNumbering", "bullet");
        CDocBuilderValue bulletNumLvl = bulletNumbering.call("GetLevel", 0);

        JSONObject workConditions = (JSONObject)data.get("work_conditions");
        document.call(
            "Push",
            createWorkCondition(api, "Working Hours", workConditions.get("working_hours").toString(), bulletNumLvl, true)
        );
        document.call(
            "Push",
            createWorkCondition(api, "Work Schedule", workConditions.get("work_schedule").toString(), bulletNumLvl, true)
        );
        JSONArray benefitsArray  = (JSONArray)workConditions.get("benefits");
        String[] benefits = new String[benefitsArray.size()];
        for (int i = 0; i < benefitsArray.size(); i++) {
            benefits[i] = benefitsArray.get(i).toString();
        }
        document.call(
            "Push",
            createWorkCondition(api, "Benefits", String.join(", ", benefits), bulletNumLvl, true)
        );
        JSONArray otherTermsArray  = (JSONArray)workConditions.get("other_terms");
        String[] otherTerms = new String[otherTermsArray.size()];
        for (int i = 0; i < otherTermsArray.size(); i++) {
            otherTerms[i] = otherTermsArray.get(i).toString();
        }
        document.call(
            "Push",
            createWorkCondition(api, "Other terms", String.join(", ", otherTerms), bulletNumLvl, false)
        );

        // TERMINATION
        document.call("Push", createNumberedSection(api, "TERMINATION", numberingLvl));
        document.call(
            "Push",
            createConditionsDescParagraph(
                api,
                String.format(
                    "Either party may terminate this Agreement by providing %s written notice. " +
                    "The Employer reserves the right to terminate employment immediately for cause, including but not limited to misconduct or breach of Agreement.",
                    ((JSONObject)data.get("termination")).get("notice_period").toString()
                )
            )
        );

        // GOVERNING LAW
        document.call("Push", createNumberedSection(api, "GOVERNING LAW", numberingLvl));
        document.call(
            "Push",
            createConditionsDescParagraph(
                api,
                String.format(
                    "This Agreement is governed by the laws of %s, and any disputes arising under this Agreement will be resolved in accordance with these laws.",
                    ((JSONObject)data.get("governing_law")).get("jurisdiction").toString()
                )
            )
        );

        // ENTIRE AGREEMENT
        document.call("Push", createNumberedSection(api, "ENTIRE AGREEMENT", numberingLvl));
        document.call(
            "Push",
            createConditionsDescParagraph(
                api,
                "This document constitutes the entire Agreement between the parties and supersedes all prior agreements. " +
                "Any amendments must be made in writing and signed by both parties."
            )
        );

        // Signatures
        CDocBuilderValue table = api.call("CreateTable", 2, 2);
        // set table properties
        table.call("SetWidth", "percent", 100);
        // fill table
        CDocBuilderValue tableTitle = table.call("GetRow", 0);
        CDocBuilderValue titleParagraph = tableTitle.call("MergeCells").call("GetContent").call("GetElement", 0);
        titleParagraph.call("Push", createRun(api, "SIGNATURES", true, 24));
        fillSigner(api, table.call("GetCell", 1, 0), "Employer");
        fillSigner(api, table.call("GetCell", 1, 1), "Employee");
        document.call("Push", table);

        // save and close
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }

    public static CDocBuilderValue createParagraph(CDocBuilderValue api, String text, boolean isBold, int fontSize, String jc) {
        CDocBuilderValue paragraph = api.call("CreateParagraph");
        paragraph.call("AddText", text);
        paragraph.call("SetBold", isBold);
        paragraph.call("SetFontSize", fontSize);
        paragraph.call("SetJc", jc);
        return paragraph;
    }

    private static CDocBuilderValue createParagraph(CDocBuilderValue api, String text, boolean isBold) {
        return createParagraph(api, text, isBold, defaultFontSize, defaultJc);
    }

    public static CDocBuilderValue createRun(CDocBuilderValue api, String text, boolean isBold, int fontSize) {
        CDocBuilderValue run = api.call("CreateRun");
        run.call("AddText", text);
        run.call("SetBold", isBold);
        run.call("SetFontSize", fontSize);
        return run;
    }

    public static void setNumbering(CDocBuilderValue paragraph, CDocBuilderValue numLvl) {
        paragraph.call("SetNumbering", numLvl);
    }

    public static void setSpacingAfter(CDocBuilderValue paragraph, int spacing) {
        paragraph.call("SetSpacingAfter", spacing);
    }

    public static CDocBuilderValue createConditionsDescParagraph(CDocBuilderValue api, String text) {
        // create paragraph with first line indentation
        CDocBuilderValue paragraph = createParagraph(api, text, false);
        paragraph.call("SetIndFirstLine", 400);
        return paragraph;
    }

    public static void addParticipantToParagraph(CDocBuilderValue api, CDocBuilderValue paragraph, String pType, String details) {
        paragraph.call("Push", createRun(api, pType + ": ", true, defaultFontSize));
        paragraph.call("Push", createRun(api, details, false, defaultFontSize));
    }

    public static CDocBuilderValue createNumberedSection(CDocBuilderValue api, String text, CDocBuilderValue numLvl) {
        CDocBuilderValue paragraph = createParagraph(api, text, true);
        setNumbering(paragraph, numLvl);
        setSpacingAfter(paragraph, 50);
        return paragraph;
    }

    public static CDocBuilderValue createWorkCondition(CDocBuilderValue api, String title, String text, CDocBuilderValue numLvl, boolean setSpacing) {
        CDocBuilderValue paragraph = api.call("CreateParagraph");
        setNumbering(paragraph, numLvl);
        if (setSpacing) {
            setSpacingAfter(paragraph, 20);
        }
        paragraph.call("SetJc", "left");
        paragraph.call("Push", createRun(api, title + ": ", true, defaultFontSize));
        paragraph.call("Push", createRun(api, text, false, defaultFontSize));
        return paragraph;
    }

    public static void fillSigner(CDocBuilderValue api, CDocBuilderValue cell, String title) {
        CDocBuilderValue paragraph = cell.call("GetContent").call("GetElement", 0);
        paragraph.call("SetJc", "left");
        paragraph.call("Push", createRun(api, title, true, defaultFontSize));

        for (String text : signerData) {
            paragraph.call("AddLineBreak");
            paragraph.call("Push", createRun(api, text, false, defaultFontSize));
        }
    }
}
