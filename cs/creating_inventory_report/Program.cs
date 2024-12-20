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
            string resultPath = "../../../result.xlsx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateInventoryReport(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateInventoryReport(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string jsonPath = resourcesDir + "/data/ims_response.json";
            string json = File.ReadAllText(jsonPath);
            InventoryData data = JsonSerializer.Deserialize<InventoryData>(json);

            // init docbuilder and create new xlsx file
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue worksheet = api.Call("GetActiveSheet");

            // fill table headers
            worksheet.Call("GetRangeByNumber", 0, 0).Call("SetValue", "Item");
            worksheet.Call("GetRangeByNumber", 0, 1).Call("SetValue", "Quantity");
            worksheet.Call("GetRangeByNumber", 0, 2).Call("SetValue", "Status");
            // make headers bold
            CValue startCell = worksheet.Call("GetRangeByNumber", 0, 0);
            CValue endCell = worksheet.Call("GetRangeByNumber", 0, 2);
            worksheet.Call("GetRange", startCell, endCell).Call("SetBold", true);
            // fill table data
            var inventory = data.inventory;
            for (int i = 0; i < inventory.Count; i++)
            {
                ItemData entry = inventory[i];
                CValue cell = worksheet.Call("GetRangeByNumber", i + 1, 0);
                cell.Call("SetValue", entry.item);
                cell = worksheet.Call("GetRangeByNumber", i + 1, 1);
                cell.Call("SetValue", entry.quantity);
                cell = worksheet.Call("GetRangeByNumber", i + 1, 2);
                string status = entry.status;
                cell.Call("SetValue", status);
                // fill cell with color corresponding to status
                if (status == "In Stock")
                    cell.Call("SetFillColor", api.Call("CreateColorFromRGB", 0, 194, 87));
                else if (status == "Reserved")
                    cell.Call("SetFillColor", api.Call("CreateColorFromRGB", 255, 255, 0));
                else
                    cell.Call("SetFillColor", api.Call("CreateColorFromRGB", 255, 79, 79));
            }
            // tweak cells width
            worksheet.Call("GetRange", "A1").Call("SetColumnWidth", 40);
            worksheet.Call("GetRange", "C1").Call("SetColumnWidth", 15);

            // save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }
    }

    // Define classes to represent the JSON structure
    public class InventoryData
    {
        public List<ItemData> inventory { get; set; }
    }

    public class ItemData
    {
        public string item { get; set; }
        public int quantity { get; set; }
        public string status { get; set; }
    }
}
