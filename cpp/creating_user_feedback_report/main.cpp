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

#include <iostream>
#include <fstream>
#include <string>
#include <algorithm>
#include <cmath>
#include <map>
#include <vector>
#include <sstream>

#include "common.h"
#include "docbuilder.h"

#include "out/cpp/builder_path.h"
#include "resources/utils/utils.h"
#include "resources/utils/json/json.hpp"

using namespace std;
using namespace NSDoctRenderer;
using json = nlohmann::json;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.xlsx";

CValue color_black;
CValue color_orange;
CValue color_grey;
CValue color_blue;

// Helper functions
string doubleToString(double value, int precision = 1) {
    std::ostringstream oss;
    oss.imbue(std::locale("en_US.UTF-8"));
    oss << std::fixed << std::setprecision(precision) << value;
    return oss.str();
}


CValue getArrayRow(const vector<string>& row_data) {
    int rowsLen = (int) row_data.size();
    CValue row = CValue::CreateArray(rowsLen);
    for (int i = 0; i < rowsLen; i++) {
        row[i] = row_data[i].c_str();
    }
    return row;
}

void setTableStyle(CValue range) {
    range.Call("SetRowHeight", 24);
    range.Call("SetAlignVertical", "center");

    string lineStyle = "Thin";
    range.Call("SetBorders", "Top", lineStyle.c_str(), color_black);
    range.Call("SetBorders", "Left", lineStyle.c_str(), color_black);
    range.Call("SetBorders", "Right", lineStyle.c_str(), color_black);
    range.Call("SetBorders", "Bottom", lineStyle.c_str(), color_black);
    range.Call("SetBorders", "InsideHorizontal", lineStyle.c_str(), color_black);
    range.Call("SetBorders", "InsideVertical", lineStyle.c_str(), color_black);
}

int fillAverageSheet(CValue worksheet, json& feedbackData) {
    map<string, vector<int>> result;
    vector<string> questionOrder;

    for (const auto& record : feedbackData) {
        for (const auto& item : record["feedback"]) {
            string question = item["question"].get<string>();
            int rating = item["answer"]["rating"].get<int>();

            if (result.find(question) == result.end()) {
                questionOrder.push_back(question);
            }
            result[question].push_back(rating);
        }
    }

    int questionSize = (int)questionOrder.size();
    CValue averageValues = CValue::CreateArray(questionSize + 1);
    averageValues[0] = getArrayRow({"Question", "Average Rating", "Number of Responses"});
    for (int i = 0; i < questionSize; i++) {
        vector<int>& ratings = result[questionOrder[i]];
        string average = doubleToString((double)accumulate(ratings.begin(), ratings.end(), 0) / ratings.size());
        averageValues[i + 1] = getArrayRow({questionOrder[i], average, to_string(ratings.size())});
    }

    int colsCount = averageValues[0].GetLength() - 1;
    int rowsCount = averageValues.GetLength();
    CValue startCell = worksheet.Call("GetRangeByNumber", 0, 0);
    CValue endCell = worksheet.Call("GetRangeByNumber", rowsCount - 1, colsCount);

    CValue averageRange = worksheet.Call("GetRange", startCell, endCell);
    setTableStyle(averageRange);
    worksheet.Call(
        "GetRange",
        worksheet.Call("GetRangeByNumber", 1, 1),
        endCell
    ).Call("SetAlignHorizontal", "center");

    CValue headerRow = worksheet.Call(
        "GetRange",
        startCell,
        worksheet.Call("GetRangeByNumber", 0, colsCount)
    );
    headerRow.Call("SetBold", true);

    averageRange.Call("SetValue", averageValues);
    averageRange.Call("AutoFit", false, true);

    return rowsCount;
}

int fillPersonalRatingsAndComments(CValue worksheet, json& feedbackData) {
    CValue headerValues = CValue::CreateArray(1);
    headerValues[0] = getArrayRow({"Date", "Question", "Comment", "Rating", "Average User Rating"});
    int colsCount = headerValues[0].GetLength() - 1;
    CValue startCell = worksheet.Call("GetRangeByNumber", 0, 0);
    CValue headerRow = worksheet.Call(
        "GetRange",
        startCell,
        worksheet.Call("GetRangeByNumber", 0, colsCount)
    );

    headerRow.Call("SetValue", headerValues);
    headerRow.Call("SetBold", true);

    int rowsCount = 1;
    for (const auto& record : feedbackData) {
        // Count and fill user feedback
        double avgRating = 0;

        int feedbackSize = (int)record["feedback"].size();
        CValue userFeedback = CValue::CreateArray(feedbackSize);
        int i = 0;
        for (const auto& item : record["feedback"]) {
            string question = item["question"].get<string>();
            string comment = item["answer"]["comment"].get<string>();
            int rating = item["answer"]["rating"].get<int>();

            userFeedback[i] = getArrayRow({question, comment, to_string(rating)});
            avgRating += rating;
            i++;
        }

        int userRowsCount = feedbackSize - 1;
        // Fill date
        CValue dateCell = worksheet.Call(
            "GetRange",
            worksheet.Call("GetRangeByNumber", rowsCount, 0),
            worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, 0)
        );
        dateCell.Call("Merge", false);
        dateCell.Call("SetValue", record["date"].get<string>().c_str());

        // Fill ratings
        CValue userRange = worksheet.Call(
            "GetRange",
            worksheet.Call("GetRangeByNumber", rowsCount, 1),
            worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, colsCount - 1)
        );
        userRange.Call("SetValue", userFeedback);

        // Count average rating
        avgRating = avgRating / feedbackSize;
        CValue ratingCell = worksheet.Call(
            "GetRange",
            worksheet.Call("GetRangeByNumber", rowsCount, colsCount),
            worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, colsCount)
        );
        ratingCell.Call("Merge", false);
        ratingCell.Call("SetValue", doubleToString(avgRating).c_str());

        // If rating <= 2, highlight it
        if (avgRating <= 2) {
            worksheet.Call(
                "GetRange",
                worksheet.Call("GetRangeByNumber", rowsCount, 0),
                worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, colsCount)
            ).Call("SetFillColor", color_orange);
        }

        // Update rows count
        rowsCount += feedbackSize;
    }

    // Format table
    rowsCount -= 1;
    CValue resultRange = worksheet.Call(
        "GetRange",
        startCell,
        worksheet.Call("GetRangeByNumber", rowsCount, colsCount)
    );
    setTableStyle(resultRange);
    worksheet.Call(
        "GetRange",
        worksheet.Call("GetRangeByNumber", 1, colsCount - 1),
        worksheet.Call("GetRangeByNumber", rowsCount, colsCount)
    ).Call("SetAlignHorizontal", "center");
    resultRange.Call("AutoFit", false, true);

    return rowsCount + 1;
}

void createColumnChart(CValue worksheet, string dataRange, string title) {
    CValue chart = worksheet.Call("AddChart", dataRange.c_str(), false, "bar", 2, 135.38 * 36000, 81.28 * 36000);
    chart.Call("SetPosition", 0, 0, 0, 0);
    chart.Call("SetTitle", title.c_str(), 16);
}

void createLineChart(CValue api, CValue worksheet, json& feedbackData, string title) {
    map<string, vector<int>> result;
    vector<string> dateOrder;

    for (const auto& record : feedbackData) {
        string date = record["date"].get<string>();
        if (result.find(date) == result.end()) {
            dateOrder.push_back(date);
        }
        for (const auto& item : record["feedback"]) {
            int rating = item["answer"]["rating"].get<int>();
            result[date].push_back(rating);
        }
    }

    int dateSize = (int)dateOrder.size();
    CValue averageDayRating = CValue::CreateArray(dateSize + 1);
    averageDayRating[0] = getArrayRow({"Date", "Rating"});
    for (int i = 0; i < dateSize; i++) {
        vector<int>& ratings = result[dateOrder[i]];
        string average = doubleToString((double)accumulate(ratings.begin(), ratings.end(), 0) / ratings.size());
        averageDayRating[i + 1] = getArrayRow({dateOrder[i], average});
    }

    string dataRange = "$E$1:$F$" + to_string(averageDayRating.GetLength());
    worksheet.Call("GetRange", dataRange.c_str()).Call("SetValue", averageDayRating);
    CValue chart = worksheet.Call("AddChart", ("Charts!" + dataRange).c_str(), false, "scatter", 2, 135.38 * 36000, 81.28 * 36000);
    chart.Call("SetPosition", 0, 0, 18, 0);
    chart.Call("SetSeriesFill", color_blue, 0, false);

    CValue stroke = api.Call(
        "CreateStroke",
        0.5 * 36000,
        api.Call("CreateSolidFill", color_grey)
    );
    chart.Call("SetSeriesOutLine", stroke, 0, false);
    chart.Call("SetTitle", title.c_str(), 16);
    chart.Call("SetMajorHorizontalGridlines", api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
}

void createPieChart(CValue api, CValue worksheet, string dataRange, string title) {
    CValue pieChartData = CValue::CreateArray(2);
    pieChartData[0] = getArrayRow({"Negative", "Neutral", "Positive"});
    pieChartData[1] = getArrayRow(
        {
            "=COUNTIF(" + dataRange + ", \"<=2\")",
            "=COUNTIF(" + dataRange + ", \"=3\")",
            "=COUNTIF(" + dataRange + ", \">=4\")"
        }
    );
    worksheet.Call("GetRange", "$A$1:$C$2").Call("SetValue", pieChartData);

    CValue chart = worksheet.Call("AddChart", "Charts!$A$1:$C$2", true, "pie", 2, 135.38 * 36000, 81.28 * 36000);
    chart.Call("SetPosition", 9, 0, 0, 0);
    chart.Call("SetTitle", title.c_str(), 16);
    chart.Call("SetDataPointFill", api.Call("CreateSolidFill", api.Call("CreateRGBColor", 237, 125, 49)), 0, 0);
    chart.Call("SetDataPointFill", api.Call("CreateSolidFill", color_grey), 0, 1);
    chart.Call("SetDataPointFill", api.Call("CreateSolidFill", color_blue), 0, 2);

    CValue stroke = api.Call(
        "CreateStroke",
        0.5 * 36000,
        api.Call("CreateSolidFill", api.Call("CreateRGBColor", 255, 255, 255))
    );
    chart.Call("SetSeriesOutLine", stroke, 0, false);
}

int main() {
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/user_feedback_data.json";
    ifstream fs(jsonPath);
    json data = json::parse(fs);

    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;
    builder.CreateFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX);

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];

    // Set main colors
    color_black = api.Call("CreateColorFromRGB", 0, 0, 0);
    color_orange = api.Call("CreateColorFromRGB", 237, 125, 49);
    color_grey = api.Call("CreateRGBColor", 128, 128, 128);
    color_blue = api.Call("CreateRGBColor", 91, 155, 213);

    // Get current worksheet
    CValue worksheet1 = api.Call("GetActiveSheet");

    // Create worksheet with average values
    worksheet1.Call("SetName", "Average");
    int table1RowsCount = fillAverageSheet(worksheet1, data);

    // Create worksheet with comments and personal ratings
    api.Call("AddSheet", "Comments");
    CValue worksheet2 = api.Call("GetActiveSheet");
    int table2RowsCount = fillPersonalRatingsAndComments(worksheet2, data);

    // Create worksheet with charts
    api.Call("AddSheet", "Charts");
    CValue worksheet3 = api.Call("GetActiveSheet");
    createColumnChart(worksheet3, "Average!$A$2:$B$" + to_string(table1RowsCount), "Average ratings");
    createLineChart(api, worksheet3, data, "Dynamics of the average ratings");
    createPieChart(api, worksheet3, "Comments!$D$1:$D$" + to_string(table2RowsCount), "Shares of reviews");

    // Set first worksheet active
    worksheet1.Call("SetActive");

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
