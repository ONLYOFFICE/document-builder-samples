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
using CContextScope = docbuilder_net.CDocBuilderContextScope;

using System;
using System.Collections.Generic;

namespace Sample
{
    public class FillingForm
    {
        public static void Main(string[] args)
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
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.OpenFile(filePath, "docxf");

            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];

            // Fill form
            CValue oDocument = oApi.Call("GetDocument");
            CValue aForms = oDocument.Call("GetAllForms");
            int formNum = 0;
            while (formNum < aForms.GetLength())
            {
                CValue form = aForms[formNum];
                string type = form.Call("GetFormType").ToString();
                string value;
                try
                {
                    value = formData[form.Call("GetFormKey").ToString()];
                }
                catch (Exception e)
                {
                    value = "";
                }
                if (type == "textForm") form.Call("SetText", value);
                if (type == "pictureForm") form.Call("SetImage", value);
                formNum++;
            }

            // Save file and close DocBuilder
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();

            CDocBuilder.Destroy();
        }
    }
}
