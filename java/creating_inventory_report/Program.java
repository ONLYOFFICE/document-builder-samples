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
    public static void main(String[] args) throws Exception {
        String resultPath = "result.xlsx";
        String resourcesDir = "../../resources";

        createInventoryReport(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createInventoryReport(String resultPath, String resourcesDir) throws Exception {
        // parse JSON
        String jsonPath = resourcesDir + "/data/ims_response.json";
        JSONObject data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));

        // init docbuilder and create new xlsx file
        int doctype = FileTypes.Spreadsheet.XLSX;
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");
        CDocBuilderValue worksheet = api.call("GetActiveSheet");

        // fill table headers
        worksheet.call("GetRangeByNumber", 0, 0).call("SetValue", "Item");
        worksheet.call("GetRangeByNumber", 0, 1).call("SetValue", "Quantity");
        worksheet.call("GetRangeByNumber", 0, 2).call("SetValue", "Status");
        // make headers bold
        CDocBuilderValue startCell = worksheet.call("GetRangeByNumber", 0, 0);
        CDocBuilderValue endCell = worksheet.call("GetRangeByNumber", 0, 2);
        worksheet.call("GetRange", startCell, endCell).call("SetBold", true);
        // fill table data
        JSONArray inventory = (JSONArray)data.get("inventory");
        for (int i = 0; i < inventory.size(); i++) {
            JSONObject entry = (JSONObject)inventory.get(i);
            CDocBuilderValue cell = worksheet.call("GetRangeByNumber", i + 1, 0);
            cell.call("SetValue", entry.get("item").toString());
            cell = worksheet.call("GetRangeByNumber", i + 1, 1);
            cell.call("SetValue", entry.get("quantity").toString());
            cell = worksheet.call("GetRangeByNumber", i + 1, 2);
            String status = entry.get("status").toString();
            cell.call("SetValue", status);
            // fill cell with color corresponding to status
            if (status.equals("In Stock"))
                cell.call("SetFillColor", api.call("CreateColorFromRGB", 0, 194, 87));
            else if (status.equals("Reserved"))
                cell.call("SetFillColor", api.call("CreateColorFromRGB", 255, 255, 0));
            else
                cell.call("SetFillColor", api.call("CreateColorFromRGB", 255, 79, 79));
        }
        // tweak cells width
        worksheet.call("GetRange", "A1").call("SetColumnWidth", 40);
        worksheet.call("GetRange", "C1").call("SetColumnWidth", 15);

        // save and close
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }
}
