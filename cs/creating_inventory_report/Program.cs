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
            string resultPath = "../../../result.xlsx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateInventoryReport(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateInventoryReport(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string json_path = resourcesDir + "/data/ims_response.json";
            string json = File.ReadAllText(json_path);
            InventoryData data = JsonSerializer.Deserialize<InventoryData>(json);

            // init docbuilder and create new xlsx file
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.CreateFile(doctype);

            CContext oContext = oBuilder.GetContext();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];
            CValue oWorksheet = oApi.Call("GetActiveSheet");

            // fill table headers
            oWorksheet.Call("GetRangeByNumber", 0, 0).Call("SetValue", "Item");
            oWorksheet.Call("GetRangeByNumber", 0, 1).Call("SetValue", "Quantity");
            oWorksheet.Call("GetRangeByNumber", 0, 2).Call("SetValue", "Status");
            // make headers bold
            CValue oStartCell = oWorksheet.Call("GetRangeByNumber", 0, 0);
            CValue oEndCell = oWorksheet.Call("GetRangeByNumber", 0, 2);
            oWorksheet.Call("GetRange", oStartCell, oEndCell).Call("SetBold", true);
            // fill table data
            var inventory = data.inventory;
            for (int i = 0; i < inventory.Count; i++)
            {
                ItemData entry = inventory[i];
                CValue oCell = oWorksheet.Call("GetRangeByNumber", i + 1, 0);
                oCell.Call("SetValue", entry.item);
                oCell = oWorksheet.Call("GetRangeByNumber", i + 1, 1);
                oCell.Call("SetValue", entry.quantity);
                oCell = oWorksheet.Call("GetRangeByNumber", i + 1, 2);
                string status = entry.status;
                oCell.Call("SetValue", status);
                // fill cell with color corresponding to status
                if (status == "In Stock")
                    oCell.Call("SetFillColor", oApi.Call("CreateColorFromRGB", 0, 194, 87));
                else if (status == "Reserved")
                    oCell.Call("SetFillColor", oApi.Call("CreateColorFromRGB", 255, 255, 0));
                else
                    oCell.Call("SetFillColor", oApi.Call("CreateColorFromRGB", 255, 79, 79));
            }
            // tweak cells width
            oWorksheet.Call("GetRange", "A1").Call("SetColumnWidth", 40);
            oWorksheet.Call("GetRange", "C1").Call("SetColumnWidth", 15);

            // save and close
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
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
