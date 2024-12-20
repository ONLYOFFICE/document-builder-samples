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
using System.Text.Json;
using System.IO;

namespace Sample
{
    public class CreatingPresentation
    {
        public static void Main(string[] args)
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.pptx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateStartupPresentation(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateStartupPresentation(string workDirectory, string resultPath, string resourcesDir)
        {
            // init docbuilder and create new pptx file
            var doctype = (int)OfficeFileTypes.Presentation.PPTX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.CreateFile(doctype);

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
            string json_path = resourcesDir + "/data/statista_api_response.json";
            string json = File.ReadAllText(json_path);
            MarketOverviewData dataMarket = JsonSerializer.Deserialize<MarketOverviewData>(json);
            // create new slide
            oSlide = addNewSlide(oApi, backgroundFill);
            // title
            oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
            addTextToParagraph(oApi, oParagraph, "Market Overview", 72, textFill, false, "center");
            // market size
            Tuple<string, string> marketSize = separateValueAndUnit(dataMarket.market.size);
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 1.58);
            addTextToParagraph(oApi, oParagraph, "Market size:", 48, textFill, false, "center");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 1.97);
            addTextToParagraph(oApi, oParagraph, marketSize.Item1, 144, textSpecialFill, false, "center", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 3.06);
            addTextToParagraph(oApi, oParagraph, marketSize.Item2, 48, textFill, false, "center");
            // growth rate
            Tuple<string, string> marketGrowth = separateValueAndUnit(dataMarket.market.growth_rate);
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 7, 1.58);
            addTextToParagraph(oApi, oParagraph, "Growth rate:", 48, textFill, false, "center");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 7, 1.97);
            addTextToParagraph(oApi, oParagraph, marketGrowth.Item1, 144, textSpecialFill, false, "center", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 7, 3.06);
            addTextToParagraph(oApi, oParagraph, marketGrowth.Item2, 48, textFill, false, "center");
            // trends
            oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 3.75);
            addTextToParagraph(oApi, oParagraph, "Trends:", 48, textFill, false, "center");
            oParagraph = addParagraphToSlide(oApi, oSlide, 0.93, 2.92, 1.57, 4.31);
            addTextToParagraph(oApi, oParagraph, makeBulletString('>', dataMarket.market.trends.Count), 72, textSpecialFill, false, "left", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 9.21, 2.92, 2.1, 4.31);
            string trendsText = "";
            foreach (string trend in dataMarket.market.trends)
            {
                trendsText += trend + "\n";
            }
            addTextToParagraph(oApi, oParagraph, trendsText, 72, textSpecialFill, false, "center", "Arial Black");

            // COMPETITORS OVERVIEW section
            // parse JSON, obtained as Statista API response
            json_path = resourcesDir + "/data/crunchbase_api_response.json";
            json = File.ReadAllText(json_path);
            CompetitorsOverviewData dataCompetitors = JsonSerializer.Deserialize<CompetitorsOverviewData>(json);
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
            int competitorsCount = dataCompetitors.competitors.Count;
            CValue[] shares = new CValue[competitorsCount + 1];
            CValue[] competitors = new CValue[competitorsCount + 1];
            for (int i = 0; i < competitorsCount; i++)
            {
                CompetitorData competitor = dataCompetitors.competitors[i];
                competitors[i] = competitor.name;
                string share = competitor.market_share;
                // remove last percent symbol
                share = share.Remove(share.Length - 1);
                othersShare -= double.Parse(share);
                shares[i] = share;
            }
            shares[competitorsCount] = othersShare.ToString().Replace(',', '.');
            competitors[competitorsCount] = "Others";
            // create a chart
            CValue arrChartData = oContext.CreateArray(1);
            arrChartData[0] = shares;
            CValue oChart = oApi.Call("CreateChart", "pie", arrChartData, oContext.CreateArray(0), competitors);
            setChartSizes(oChart, 6.51, 5.9, 4.18, 1.49);
            oChart.Call("SetLegendFontSize", 14);
            oChart.Call("SetLegendPos", "right");
            oSlide.Call("AddObject", oChart);

            // create slide for every competitor with brief info
            foreach (CompetitorData competitor in dataCompetitors.competitors)
            {
                // create new slide
                oSlide = addNewSlide(oApi, backgroundFill);
                // title
                oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
                addTextToParagraph(oApi, oParagraph, "Competitors Overview", 72, textFill, false, "center");
                // header
                oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 1.2);
                addTextToParagraph(oApi, oParagraph, competitor.name, 64, textFill, false, "center");
                // recent funding
                oParagraph = addParagraphToSlide(oApi, oSlide, 3.13, 0.8, 1.07, 2.65);
                addTextToParagraph(oApi, oParagraph, "Recent funding:", 48, textFill);
                oParagraph = addParagraphToSlide(oApi, oSlide, 8.9, 0.8, 4.19, 2.52);
                addTextToParagraph(oApi, oParagraph, competitor.recent_funding, 96, textSpecialFill, false, "left", "Arial Black");
                // main products
                oParagraph = addParagraphToSlide(oApi, oSlide, 3.13, 0.8, 1.07, 3.72);
                addTextToParagraph(oApi, oParagraph, "Main products:", 48, textFill);
                oParagraph = addParagraphToSlide(oApi, oSlide, 0.93, 3.53, 4.19, 3.72);
                addTextToParagraph(oApi, oParagraph, makeBulletString('>', competitor.products.Count), 72, textSpecialFill, false, "left", "Arial Black");
                oParagraph = addParagraphToSlide(oApi, oSlide, 7.97, 3.53, 5.12, 3.72);
                string productsText = "";
                foreach (string product in competitor.products)
                {
                    productsText += product + '\n';
                }
                addTextToParagraph(oApi, oParagraph, productsText, 72, textSpecialFill, false, "left", "Arial Black");
            }

            // TARGET AUDIENCE section
            // parse JSON, obtained as Social Media Insights API response
            json_path = resourcesDir + "/data/smi_api_response.json";
            json = File.ReadAllText(json_path);
            AudienceOverviewData dataAudience = JsonSerializer.Deserialize<AudienceOverviewData>(json);
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
            addTextToParagraph(oApi, oParagraph, dataAudience.demographics.age_range, 128, textSpecialFill, false, "center", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 2.95);
            addTextToParagraph(oApi, oParagraph, "age range", 40, textFill, false, "center");
            // location
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 3.68);
            addTextToParagraph(oApi, oParagraph, dataAudience.demographics.location, 72, textSpecialFill, false, "center", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 4.27);
            addTextToParagraph(oApi, oParagraph, "location", 40, textFill, false, "center");
            // income level
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 5.28);
            addTextToParagraph(oApi, oParagraph, dataAudience.demographics.income_level, 56, textSpecialFill, false, "center", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 0.8, 5.83);
            addTextToParagraph(oApi, oParagraph, "income level", 40, textFill, false, "center");

            // social trends
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 6, 1.33);
            addTextToParagraph(oApi, oParagraph, "Social trends:", 48, textFill, false, "center");
            // positive feedback
            oParagraph = addParagraphToSlide(oApi, oSlide, 0.63, 2.42, 7, 2.06);
            addTextToParagraph(oApi, oParagraph, makeBulletString('+', dataAudience.social_trends.positive_feedback.Count), 52, textSpecialFill, false, "left", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.56, 2.42, 7.67, 2.06);
            string positiveFeedback = "";
            foreach (string feedback in dataAudience.social_trends.positive_feedback)
            {
                positiveFeedback += feedback + '\n';
            }
            addTextToParagraph(oApi, oParagraph, positiveFeedback, 52, textSpecialFill, false, "left", "Arial Black");
            // negative feedback
            oParagraph = addParagraphToSlide(oApi, oSlide, 0.63, 2.42, 7, 4.55);
            addTextToParagraph(oApi, oParagraph, makeBulletString('-', dataAudience.social_trends.negative_feedback.Count), 52, textAltFill, false, "left", "Arial Black");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.56, 2.42, 7.67, 4.55);
            string negativeFeedback = "";
            foreach (string feedback in dataAudience.social_trends.negative_feedback)
            {
                negativeFeedback += feedback + '\n';
            }
            addTextToParagraph(oApi, oParagraph, negativeFeedback, 52, textAltFill, false, "left", "Arial Black");

            // SEARCH TRENDS section
            // parse JSON, obtained as Google Trends API response
            json_path = resourcesDir + "/data/google_trends_api_response.json";
            json = File.ReadAllText(json_path);
            TrendsOverviewData dataTrends = JsonSerializer.Deserialize<TrendsOverviewData>(json);
            // create new slide
            oSlide = addNewSlide(oApi, backgroundFill);
            // title
            oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
            addTextToParagraph(oApi, oParagraph, "Search Trends", 72, textFill, false, "center");
            // add every trend on the slide
            double offsetY = 1.43;
            foreach (SearchTrendData trend in dataTrends.search_trends)
            {
                oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, offsetY);
                addTextToParagraph(oApi, oParagraph, trend.topic, 96, textSpecialFill, false, "center", "Arial Black");
                oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, offsetY + 0.8);
                addTextToParagraph(oApi, oParagraph, trend.growth, 40, textFill, false, "center");
                offsetY += 1.25;
            }

            // FINANCIAL MODEL section
            // parse JSON, obtained from financial system
            json_path = resourcesDir + "/data/financial_model_data.json";
            json = File.ReadAllText(json_path);
            FinancialModelData dataFinancial = JsonSerializer.Deserialize<FinancialModelData>(json);
            // create new slide
            oSlide = addNewSlide(oApi, backgroundFill);
            // title
            oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 0.4);
            addTextToParagraph(oApi, oParagraph, "Financial Model", 72, textFill, false, "center");
            // chart title
            oParagraph = addParagraphToSlide(oApi, oSlide, 11.8, 0.8, 0.8, 1.2);
            addTextToParagraph(oApi, oParagraph, "Profit forecast", 48, textFill, false, "center");
            // chart
            string[] chartKeys = { "revenue", "cost_of_goods_sold", "gross_profit", "operating_expenses", "net_profit" };
            var profitForecast = dataFinancial.profit_forecast;
            CValue arrChartYears = oContext.CreateArray(profitForecast.Count);
            for (int i = 0; i < profitForecast.Count; i++)
            {
                arrChartYears[i] = profitForecast[i].year;
            }
            arrChartData = oContext.CreateArray(chartKeys.Length);
            for (int i = 0; i < chartKeys.Length; i++)
            {
                arrChartData[i] = oContext.CreateArray(profitForecast.Count);
                for (int j = 0; j < profitForecast.Count; j++)
                {
                    arrChartData[i][j] = (string)profitForecast[j].GetType().GetProperty(chartKeys[i]).GetValue(profitForecast[j]);
                }
            }
            CValue arrChartNames = new CValue[] { "Revenue", "Cost of goods sold", "Gross profit", "Operating expenses", "Net profit" };
            oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, arrChartNames, arrChartYears);
            setChartSizes(oChart, 10.06, 5.06, 1.67, 2);
            string moneyUnit = separateValueAndUnit(profitForecast[0].revenue).Item2;
            string verAxisTitle = "Amount (" + moneyUnit + ")";
            oChart.Call("SetVerAxisTitle", verAxisTitle, 14, false);
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
            Tuple<string, string> fixedCostsWithUnit = separateValueAndUnit(dataFinancial.break_even_analysis.fixed_costs);
            double fixedCosts = double.Parse(fixedCostsWithUnit.Item1);
            moneyUnit = fixedCostsWithUnit.Item2;
            double sellingPricePerUnit = double.Parse(separateValueAndUnit(dataFinancial.break_even_analysis.selling_price_per_unit).Item1);
            double variableCostPerUnit = double.Parse(separateValueAndUnit(dataFinancial.break_even_analysis.variable_cost_per_unit).Item1);
            int breakEvenPoint = dataFinancial.break_even_analysis.break_even_point;
            int step = breakEvenPoint / 4;
            CValue chartUnits = oContext.CreateArray(9);
            CValue chartRevenue = oContext.CreateArray(9);
            CValue chartTotalCosts = oContext.CreateArray(9);
            for (int i = 0; i < 9; i++)
            {
                int currUnits = i * step;
                chartUnits[i] = currUnits;
                chartRevenue[i] = currUnits * sellingPricePerUnit;
                chartTotalCosts[i] = fixedCosts + currUnits * variableCostPerUnit;
            }
            arrChartData = new CValue[] { chartRevenue, chartTotalCosts };
            arrChartNames = new CValue[] { "Revenue", "Total costs" };
            // create chart
            oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, arrChartNames, chartUnits);
            setChartSizes(oChart, 9.17, 5.06, 0.31, 2);
            verAxisTitle = "Amount (" + moneyUnit + ")";
            oChart.Call("SetVerAxisTitle", verAxisTitle, 14, false);
            oChart.Call("SetHorAxisTitle", "Units sold", 14, false);
            oChart.Call("SetLegendFontSize", 14);
            oChart.Call("SetMinorVerticalGridlines", stroke);
            oChart.Call("SetLegendPos", "bottom");
            oSlide.Call("AddObject", oChart);
            // break even point
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 8.4, 3.11);
            addTextToParagraph(oApi, oParagraph, "Break even point:", 48, textFill, false, "center");
            oParagraph = addParagraphToSlide(oApi, oSlide, 5.62, 0.8, 8.4, 3.51);
            addTextToParagraph(oApi, oParagraph, breakEvenPoint.ToString(), 128, textSpecialFill, false, "center", "Arial Black");
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
            var growthRates = dataFinancial.growth_rates;
            arrChartYears = oContext.CreateArray(growthRates.Count);
            CValue arrChartGrowth = oContext.CreateArray(growthRates.Count);
            for (int i = 0; i < growthRates.Count; i++)
            {
                arrChartYears[i] = growthRates[i].year;
                arrChartGrowth[i] = growthRates[i].growth;
            }
            arrChartData = new CValue[] { arrChartGrowth };
            oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, oContext.CreateArray(0), arrChartYears);
            setChartSizes(oChart, 10.06, 5.06, 1.67, 2);
            oChart.Call("SetVerAxisTitle", "Growth (%)", 14, false);
            oChart.Call("SetHorAxisTitle", "Year", 14, false);
            oChart.Call("SetMinorVerticalGridlines", stroke);
            oSlide.Call("AddObject", oChart);

            // save and close
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void addTextToParagraph(CValue oApi, CValue oParagraph, string text, int fontSize, CValue oFill, bool isBold = false, string jc = "left", string fontFamily = "Arial")
        {
            CValue oRun = oApi.Call("CreateRun");
            oRun.Call("AddText", text);
            oRun.Call("SetFontSize", fontSize);
            oRun.Call("SetBold", isBold);
            oRun.Call("SetFill", oFill);
            oRun.Call("SetFontFamily", fontFamily);
            oParagraph.Call("AddElement", oRun);
            oParagraph.Call("SetJc", jc);
        }

        public static CValue addNewSlide(CValue oApi, CValue oFill)
        {
            CValue oSlide = oApi.Call("CreateSlide");
            CValue oPresentation = oApi.Call("GetPresentation");
            oPresentation.Call("AddSlide", oSlide);
            oSlide.Call("SetBackground", oFill);
            oSlide.Call("RemoveAllObjects");
            return oSlide;
        }

        public const int em_in_inch = 914400;
        // width, height, pos_x and pos_y are set in INCHES
        public static CValue addParagraphToSlide(CValue oApi, CValue oSlide, double width, double height, double pos_x, double pos_y)
        {
            CValue oShape = oApi.Call("CreateShape", "rect", width * em_in_inch, height * em_in_inch);
            oShape.Call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
            CValue oParagraph = oShape.Call("GetDocContent").Call("GetElement", 0);
            oSlide.Call("AddObject", oShape);
            return oParagraph;
        }

        public static void setChartSizes(CValue oChart, double width, double height, double pos_x, double pos_y)
        {
            oChart.Call("SetSize", width * em_in_inch, height * em_in_inch);
            oChart.Call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
        }

        public static Tuple<string, string> separateValueAndUnit(string data)
        {
            int spacePos = data.IndexOf(' ');
            return Tuple.Create(data.Substring(0, spacePos), data.Substring(spacePos + 1));
        }

        public static string makeBulletString(char bullet, int repeats)
        {
            string result = "";
            for (int i = 0; i < repeats; i++)
            {
                result += bullet;
                result += '\n';
            }
            return result;
        }
    }

    // Define classes to represent the JSON structure
    public class MarketOverviewData
    {
        public MarketData market { get; set; }
}

    public class MarketData
    {
        public string size { get; set; }
        public string growth_rate { get; set; }
        public List<string> trends { get; set; }
    }

    public class CompetitorsOverviewData
    {
        public List<CompetitorData> competitors { get; set; }
    }

    public class CompetitorData
    {
        public string name { get; set; }
        public string market_share { get; set; }
        public string recent_funding { get; set; }
        public List<string> products { get; set; }
    }

    public class AudienceOverviewData
    {
        public SocialTrendsData social_trends { get; set; }
        public DemographicsData demographics { get; set; }
    }

    public class SocialTrendsData
    {
        public List<string> positive_feedback { get; set; }
        public List<string> negative_feedback { get; set; }
    }

    public class DemographicsData
    {
        public string age_range { get; set; }
        public string location { get; set; }
        public string income_level { get; set; }
    }

    public class TrendsOverviewData
    {
        public List<SearchTrendData> search_trends { get; set; }
    }

    public class SearchTrendData
    {
        public string topic { get; set; }
        public string growth { get; set; }
    }

    public class FinancialModelData
    {
        public List<YearForecastData> profit_forecast { get; set; }
        public BreakEvenAnalysisData break_even_analysis { get; set; }
        public List<YearGrowthData> growth_rates { get; set; }
    }

    public class YearForecastData
    {
        public string year { get; set; }
        public string revenue { get; set; }
        public string cost_of_goods_sold { get; set; }
        public string gross_profit { get; set; }
        public string operating_expenses { get; set; }
        public string net_profit { get; set; }
    }

    public class BreakEvenAnalysisData
    {
        public string fixed_costs { get; set; }
        public string selling_price_per_unit { get; set; }
        public string variable_cost_per_unit { get; set; }
        public int break_even_point { get; set; }
    }

    public class YearGrowthData
    {
        public string year { get; set; }
        public string growth { get; set; }
    }
}
