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

using System;
using System.Collections.Generic;

namespace Sample
{
    public class FillingForm
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string filePath = "../../../../../../resources/docs/form.docx";
            string resultPath = "../../../result.docx";

            IDictionary<string, string> formData = new Dictionary<string, string>() {
                { "Photo", "https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png" },
                { "Serial","A1345" },
                { "Company Name", "ONLYOFFICE" },
                { "Date", "25.12.2023" },
                { "Recipient", "Space Corporation" },
                { "Qty1", "25" },
                { "Description1", "Frame" },
                { "Qty2", "2" },
                { "Description2", "Stack" },
                { "Qty3", "34" },
                { "Description3", "Shifter" }
            };
            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            FillForm(workDirectory, resultPath, filePath, formData);
        }

        public static void FillForm(string workDirectory, string resultPath, string filePath, IDictionary<string, string> formData)
        {
            var doctype = (int)OfficeFileTypes.Document.DOCX;

            // Init DocBuilder
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.OpenFile(filePath, "docxf");

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];

            // Fill form
            CValue document = api.Call("GetDocument");
            CValue forms = document.Call("GetAllForms");
            int formNum = 0;
            while (formNum < forms.GetLength())
            {
                CValue form = forms[formNum];
                string type = form.Call("GetFormType").ToString();
                string value;
                try
                {
                    value = formData[form.Call("GetFormKey").ToString()];
                }
                catch (Exception)
                {
                    value = "";
                }
                if (type == "textForm") form.Call("SetText", value);
                if (type == "pictureForm") form.Call("SetImage", value);
                formNum++;
            }

            // Save file and close DocBuilder
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();

            CDocBuilder.Destroy();
        }
    }
}
