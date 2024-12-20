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

#include <fstream>
#include <string>
#include <vector>

#include "common.h"
#include "docbuilder.h"

#include "out/cpp/builder_path.h"
#include "resources/utils/utils.h"
#include "resources/utils/json/json.hpp"

using namespace std;
using namespace NSDoctRenderer;
using json = nlohmann::json;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.pptx";

// Helper functions
void addTextToParagraph(CValue oApi, CValue oParagraph, const string& text, int fontSize, CValue oFill, bool isBold = false, string jc = "left", string fontFamily = "Arial")
{
    CValue oRun = oApi.Call("CreateRun");
    oRun.Call("AddText", text.c_str());
    oRun.Call("SetFontSize", fontSize);
    oRun.Call("SetBold", isBold);
    oRun.Call("SetFill", oFill);
    oRun.Call("SetFontFamily", fontFamily.c_str());
    oParagraph.Call("AddElement", oRun);
    oParagraph.Call("SetJc", jc.c_str());
}

CValue addNewSlide(CValue oApi, CValue oFill)
{
    CValue oSlide = oApi.Call("CreateSlide");
    CValue oPresentation = oApi.Call("GetPresentation");
    oPresentation.Call("AddSlide", oSlide);
    oSlide.Call("SetBackground", oFill);
    oSlide.Call("RemoveAllObjects");
    return oSlide;
}

constexpr int em_in_inch = 914400;
// width, height, pos_x and pos_y are set in INCHES
CValue addParagraphToSlide(CValue oApi, CValue oSlide, double width, double height, double pos_x, double pos_y)
{
    CValue oShape = oApi.Call("CreateShape", "rect", width * em_in_inch, height * em_in_inch);
    oShape.Call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
    CValue oParagraph = oShape.Call("GetDocContent").Call("GetElement", 0);
    oSlide.Call("AddObject", oShape);
    return oParagraph;
}

void setChartSizes(CValue oChart, double width, double height, double pos_x, double pos_y)
{
    oChart.Call("SetSize", width * em_in_inch, height * em_in_inch);
    oChart.Call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
}

pair<string, string> separateValueAndUnit(const string& data)
{
    size_t spacePos = data.find(' ');
    return { data.substr(0, spacePos), data.substr(spacePos + 1) };
}

string makeBulletString(char bullet, int repeats)
{
    string result;
    for (int i = 0; i < repeats; i++)
    {
        result += bullet;
        result += "\n";
    }
    return result;
}

CValue createStringArray(const vector<string>& values)
{
    CValue arrResult = CValue::CreateArray((int)values.size());
    for (int i = 0; i < values.size(); i++)
    {
        arrResult[i] = values[i].c_str();
    }

    return arrResult;
}

// Main function
int main()
{
    string resourcesDir = U_TO_UTF8(NSUtils::GetResourcesDirectory());

    // init docbuilder and create new pptx file
    CDocBuilder::Initialize(workDir);
    CDocBuilder oBuilder;
    oBuilder.CreateFile(OFFICESTUDIO_FILE_PRESENTATION_PPTX);

    CContext oContext = oBuilder.GetContext();
    CValue oGlobal = oContext.GetGlobal();
    CValue oApi = oGlobal["Api"];
    CValue oPresentation = oApi.Call("GetPresentation");

    // init colors
    CValue backgroundFill = oApi.Call("CreateSolidFill", oApi.Call("CreateRGBColor", 255, 255, 255));
    CValue textFill = oApi.Call("CreateSolidFill", oApi.Call("CreateRGBColor", 80, 80, 80));
    CValue textSpecialFill = oApi.Call("CreateSolidFill", oApi.Call("CreateRGBColor", 15, 102, 7));
    CValue textAltFill = oApi.Call("CreateSolidFill", oApi.Call("CreateRGBColor", 230, 69, 69));
    CValue chartGridFill = oApi.Call("CreateSolidFill", oApi.Call("CreateRGBColor", 134, 134, 134));
    CValue master = oPresentation.Call("GetMaster", 0);
    CValue colorScheme = master.Call("GetTheme").Call("GetColorScheme");
    colorScheme.Call("ChangeColor", 0, oApi.Call("CreateRGBColor", 15, 102, 7));

    // TITLE slide
    CValue oSlide = oPresentation.Call("GetSlideByIndex", 0);
    oSlide.Call("SetBackground", backgroundFill);
    CValue oParagraph = oSlide.Call("GetAllShapes")[0].Call("GetContent").Call("GetElement", 0);
    addTextToParagraph(oApi, oParagraph, "GreenVibe Solutions", 120, textSpecialFill, true, "center", "Arial Black");
    oParagraph = oSlide.Call("GetAllShapes")[1].Call("GetContent").Call("GetElement", 0);
    addTextToParagraph(oApi, oParagraph, "12.12.2024", 48, textFill, false, "center");

    // MARKET OVERVIEW slide
    // parse JSON, obtained as Statista API response
    ifstream fs(resourcesDir + "/data/statista_api_response.json");
    json data = json::parse(fs);
    // create new slide
    oSlide = addNewSlide(oApi, backgroundFill);
    // title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
    addTextToParagraph(oApi, oParagraph, "Market Overview", 72, textFill, false, "center");
    // market size
    pair<string, string> marketSize = separateValueAndUnit(data["market"]["size"].get<string>());
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 1.58);
    addTextToParagraph(oApi, oParagraph, "Market size:", 48, textFill, false, "center");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 1.97);
    addTextToParagraph(oApi, oParagraph, marketSize.first, 144, textSpecialFill, false, "center", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 3.06);
    addTextToParagraph(oApi, oParagraph, marketSize.second, 48, textFill, false, "center");
    // growth rate
    pair<string, string> marketGrowth = separateValueAndUnit(data["market"]["growth_rate"].get<string>());
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 7, 1.58);
    addTextToParagraph(oApi, oParagraph, "Growth rate:", 48, textFill, false, "center");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 7, 1.97);
    addTextToParagraph(oApi, oParagraph, marketGrowth.first, 144, textSpecialFill, false, "center", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 7, 3.06);
    addTextToParagraph(oApi, oParagraph, marketGrowth.second, 48, textFill, false, "center");
    // trends
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 3.75);
    addTextToParagraph(oApi, oParagraph, "Trends:", 48, textFill, false, "center");
    oParagraph = addParagraphToSlide(oApi, oSlide, 0.93, 2.92, 1.57, 4.31);
    addTextToParagraph(oApi, oParagraph, makeBulletString('>', (int)data["market"]["trends"].size()), 72, textSpecialFill, false, "left", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 9.21, 2.92, 2.1, 4.31);
    string trendsText = "";
    for (const auto& trend : data["market"]["trends"])
    {
        trendsText += trend.get<string>() + "\n";
    }
    addTextToParagraph(oApi, oParagraph, trendsText, 72, textSpecialFill, false, "center", "Arial Black");

    // COMPETITORS OVERVIEW section
    // parse JSON, obtained as Statista API response
    fs.close();
    fs.open(resourcesDir + "/data/crunchbase_api_response.json");
    data = json::parse(fs);
    // create new slide
    oSlide = addNewSlide(oApi, backgroundFill);
    // title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
    addTextToParagraph(oApi, oParagraph, "Competitors Overview", 72, textFill, false, "center");
    // chart header
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 1.2);
    addTextToParagraph(oApi, oParagraph, "Market shares", 48, textFill, false, "center");
    // get chart data
    double othersShare = 100.0;
    vector<string> shares;
    vector<string> competitors;
    for (const auto& competitor : data["competitors"])
    {
        competitors.push_back(competitor["name"].get<string>());
        string share = competitor["market_share"].get<string>();
        // remove last percent symbol
        share.pop_back();
        othersShare -= stod(share);
        shares.push_back(share);
    }
    shares.push_back(to_string(othersShare));
    competitors.push_back("Others");
    // create a chart
    CValue arrChartData = CValue::CreateArray(1);
    arrChartData[0] = createStringArray(shares);
    CValue oChart = oApi.Call("CreateChart", "pie", arrChartData, CValue::CreateArray(0), createStringArray(competitors));
    setChartSizes(oChart, 6.51, 5.9, 4.18, 1.49);
    oChart.Call("SetLegendFontSize", 14);
    oChart.Call("SetLegendPos", "right");
    oSlide.Call("AddObject", oChart);

    // create slide for every competitor with brief info
    for (const auto& competitor : data["competitors"])
    {
        // create new slide
        oSlide = addNewSlide(oApi, backgroundFill);
        // title
        oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(oApi, oParagraph, "Competitors Overview", 72, textFill, false, "center");
        // header
        oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 1.2);
        addTextToParagraph(oApi, oParagraph, competitor["name"].get<string>(), 64, textFill, false, "center");
        // recent funding
        oParagraph = addParagraphToSlide(oApi, oSlide, 3.13, 0.8, 1.07, 2.65);
        addTextToParagraph(oApi, oParagraph, "Recent funding:", 48, textFill);
        oParagraph = addParagraphToSlide(oApi, oSlide, 8.9, 0.8, 4.19, 2.52);
        addTextToParagraph(oApi, oParagraph, competitor["recent_funding"].get<string>(), 96, textSpecialFill, false, "left", "Arial Black");
        // main products
        oParagraph = addParagraphToSlide(oApi, oSlide, 3.13, 0.8, 1.07, 3.72);
        addTextToParagraph(oApi, oParagraph, "Main products:", 48, textFill);
        oParagraph = addParagraphToSlide(oApi, oSlide, 0.93, 3.53, 4.19, 3.72);
        addTextToParagraph(oApi, oParagraph, makeBulletString('>', (int)competitor["products"].size()), 72, textSpecialFill, false, "left", "Arial Black");
        oParagraph = addParagraphToSlide(oApi, oSlide, 7.97, 3.53, 5.12, 3.72);
        string productsText;
        for (const auto& product : competitor["products"])
        {
            productsText += product.get<string>() + "\n";
        }
        addTextToParagraph(oApi, oParagraph, productsText, 72, textSpecialFill, false, "left", "Arial Black");
    }

    // TARGET AUDIENCE section
    // parse JSON, obtained as Social Media Insights API response
    fs.close();
    fs.open(resourcesDir + "/data/smi_api_response.json");
    data = json::parse(fs);
    // create new slide
    oSlide = addNewSlide(oApi, backgroundFill);
    // title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
    addTextToParagraph(oApi, oParagraph, "Target Audience", 72, textFill, false, "center");

    // demographics
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 1.33);
    addTextToParagraph(oApi, oParagraph, "Demographics:", 48, textFill, false, "center");
    // age range
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 1.97);
    addTextToParagraph(oApi, oParagraph, data["demographics"]["age_range"].get<string>(), 128, textSpecialFill, false, "center", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 2.95);
    addTextToParagraph(oApi, oParagraph, "age range", 40, textFill, false, "center");
    // location
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 3.68);
    addTextToParagraph(oApi, oParagraph, data["demographics"]["location"].get<string>(), 72, textSpecialFill, false, "center", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 4.27);
    addTextToParagraph(oApi, oParagraph, "location", 40, textFill, false, "center");
    // income level
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 5.28);
    addTextToParagraph(oApi, oParagraph, data["demographics"]["income_level"].get<string>(), 56, textSpecialFill, false, "center", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 5.83);
    addTextToParagraph(oApi, oParagraph, "income level", 40, textFill, false, "center");

    // social trends
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 6, 1.33);
    addTextToParagraph(oApi, oParagraph, "Social trends:", 48, textFill, false, "center");
    // positive feedback
    oParagraph = addParagraphToSlide(oApi, oSlide, 0.63, 2.42, 7, 2.06);
    addTextToParagraph(oApi, oParagraph, makeBulletString('+', (int)data["social_trends"]["positive_feedback"].size()), 52, textSpecialFill, false, "left", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.56, 2.42, 7.67, 2.06);
    string positiveFeedback;
    for (const auto& feedback : data["social_trends"]["positive_feedback"])
    {
        positiveFeedback += feedback.get<string>() + '\n';
    }
    addTextToParagraph(oApi, oParagraph, positiveFeedback, 52, textSpecialFill, false, "left", "Arial Black");
    // negative feedback
    oParagraph = addParagraphToSlide(oApi, oSlide, 0.63, 2.42, 7, 4.55);
    addTextToParagraph(oApi, oParagraph, makeBulletString('-', (int)data["social_trends"]["negative_feedback"].size()), 52, textAltFill, false, "left", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.56, 2.42, 7.67, 4.55);
    string negativeFeedback;
    for (const auto& feedback : data["social_trends"]["negative_feedback"])
    {
        negativeFeedback += feedback.get<string>() + '\n';
    }
    addTextToParagraph(oApi, oParagraph, negativeFeedback, 52, textAltFill, false, "left", "Arial Black");

    // SEARCH TRENDS section
    // parse JSON, obtained as Google Trends API response
    fs.close();
    fs.open(resourcesDir + "/data/google_trends_api_response.json");
    data = json::parse(fs);
    // create new slide
    oSlide = addNewSlide(oApi, backgroundFill);
    // title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
    addTextToParagraph(oApi, oParagraph, "Search Trends", 72, textFill, false, "center");
    // add every trend on the slide
    double offsetY = 1.43;
    for (const auto& trend : data["search_trends"])
    {
        oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, offsetY);
        addTextToParagraph(oApi, oParagraph, trend["topic"].get<string>(), 96, textSpecialFill, false, "center", "Arial Black");
        oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, offsetY + 0.8);
        addTextToParagraph(oApi, oParagraph, trend["growth"].get<string>(), 40, textFill, false, "center");
        offsetY += 1.25;
    }

    // FINANCIAL MODEL section
    // parse JSON, obtained from financial system
    fs.close();
    fs.open(resourcesDir + "/data/financial_model_data.json");
    data = json::parse(fs);
    // create new slide
    oSlide = addNewSlide(oApi, backgroundFill);
    // title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
    addTextToParagraph(oApi, oParagraph, "Financial Model", 72, textFill, false, "center");
    // chart title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 1.2);
    addTextToParagraph(oApi, oParagraph, "Profit forecast", 48, textFill, false, "center");
    // chart
    vector<string> chartKeys = { "revenue", "cost_of_goods_sold", "gross_profit", "operating_expenses", "net_profit" };
    const json& profitForecast = data["profit_forecast"];
    CValue arrChartYears = CValue::CreateArray((int)profitForecast.size());
    for (int i = 0; i < profitForecast.size(); i++)
    {
        arrChartYears[i] = profitForecast[i]["year"].get<string>().c_str();
    }
    arrChartData = CValue::CreateArray((int)chartKeys.size());
    for (int i = 0; i < chartKeys.size(); i++)
    {
        arrChartData[i] = CValue::CreateArray((int)profitForecast.size());
        for (int j = 0; j < profitForecast.size(); j++)
        {
            arrChartData[i][j] = profitForecast[j][chartKeys[i]].get<string>().c_str();
        }
    }
    CValue arrChartNames = createStringArray({ "Revenue", "Cost of goods sold", "Gross profit", "Operating expenses", "Net profit" });
    oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, arrChartNames, arrChartYears);
    setChartSizes(oChart, 10.06, 5.06, 1.67, 2);
    string moneyUnit = separateValueAndUnit(profitForecast[0]["revenue"].get<string>()).second;
    string verAxisTitle = "Amount (" + moneyUnit + ")";
    oChart.Call("SetVerAxisTitle", verAxisTitle.c_str(), 14, false);
    oChart.Call("SetHorAxisTitle", "Year", 14, false);
    oChart.Call("SetLegendFontSize", 14);
    CValue stroke = oApi.Call("CreateStroke", 1, chartGridFill);
    oChart.Call("SetMinorVerticalGridlines", stroke);
    oSlide.Call("AddObject", oChart);

    // break even analysis
    // create new slide
    oSlide = addNewSlide(oApi, backgroundFill);
    // title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
    addTextToParagraph(oApi, oParagraph, "Financial Model", 72, textFill, false, "center");
    // chart title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 1.2);
    addTextToParagraph(oApi, oParagraph, "Break even analysis", 48, textFill, false, "center");
    // chart data
    pair<string, string> fixedCostsWithUnit = separateValueAndUnit(data["break_even_analysis"]["fixed_costs"].get<string>());
    double fixedCosts = stod(fixedCostsWithUnit.first);
    moneyUnit = fixedCostsWithUnit.second;
    double sellingPricePerUnit = stod(separateValueAndUnit(data["break_even_analysis"]["selling_price_per_unit"].get<string>()).first);
    double variableCostPerUnit = stod(separateValueAndUnit(data["break_even_analysis"]["variable_cost_per_unit"].get<string>()).first);
    int breakEvenPoint = data["break_even_analysis"]["break_even_point"].get<int>();
    int step = breakEvenPoint / 4;
    CValue chartUnits = oContext.CreateArray(9);
    CValue chartRevenue = oContext.CreateArray(9);
    CValue chartTotalCosts = oContext.CreateArray(9);
    for (int i = 0; i < 9; i++)
    {
        int currUnits = i * step;
        chartUnits[i] = currUnits;
        chartRevenue[i] = (int)(currUnits * sellingPricePerUnit);
        chartTotalCosts[i] = (int)(fixedCosts + currUnits * variableCostPerUnit);
    }
    arrChartData = CValue::CreateArray(2);
    arrChartData[0] = chartRevenue;
    arrChartData[1] = chartTotalCosts;
    arrChartNames = createStringArray({ "Revenue", "Total costs" });
    // create chart
    oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, arrChartNames, chartUnits);
    setChartSizes(oChart, 9.17, 5.06, 0.31, 2);
    verAxisTitle = "Amount (" + moneyUnit + ")";
    oChart.Call("SetVerAxisTitle", verAxisTitle.c_str(), 14, false);
    oChart.Call("SetHorAxisTitle", "Units sold", 14, false);
    oChart.Call("SetLegendFontSize", 14);
    oChart.Call("SetMinorVerticalGridlines", stroke);
    oChart.Call("SetLegendPos", "bottom");
    oSlide.Call("AddObject", oChart);
    // break even point
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 8.4, 3.11);
    addTextToParagraph(oApi, oParagraph, "Break even point:", 48, textFill, false, "center");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 8.4, 3.51);
    addTextToParagraph(oApi, oParagraph, to_string(breakEvenPoint), 128, textSpecialFill, false, "center", "Arial Black");
    oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 8.4, 4.38);
    addTextToParagraph(oApi, oParagraph, "units", 40, textFill, false, "center");

    // growth rates
    // create new slide
    oSlide = addNewSlide(oApi, backgroundFill);
    // title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
    addTextToParagraph(oApi, oParagraph, "Financial Model", 72, textFill, false, "center");
    // chart title
    oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 1.2);
    addTextToParagraph(oApi, oParagraph, "Growth rates", 48, textFill, false, "center");
    // chart
    const json& growthRates = data["growth_rates"];
    arrChartYears = CValue::CreateArray((int)growthRates.size());
    CValue arrChartGrowth = CValue::CreateArray((int)growthRates.size());
    for (int i = 0; i < growthRates.size(); i++)
    {
        arrChartYears[i] = growthRates[i]["year"].get<string>().c_str();
        arrChartGrowth[i] = growthRates[i]["growth"].get<string>().c_str();
    }
    arrChartData = CValue::CreateArray(1);
    arrChartData[0] = arrChartGrowth;
    oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, CValue::CreateArray(0), arrChartYears);
    setChartSizes(oChart, 10.06, 5.06, 1.67, 2);
    oChart.Call("SetVerAxisTitle", "Growth (%)", 14, false);
    oChart.Call("SetHorAxisTitle", "Year", 14, false);
    oChart.Call("SetMinorVerticalGridlines", stroke);
    oSlide.Call("AddObject", oChart);

    // save and close
    oBuilder.SaveFile(OFFICESTUDIO_FILE_PRESENTATION_PPTX, resultPath);
    oBuilder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
