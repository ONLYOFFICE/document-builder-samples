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

import docbuilder.*;

import java.io.File;
import java.io.FileReader;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class Program {
    public static void main(String[] args) throws Exception {
        String resultPath = "result.pptx";
        String resourcesDir = "../../resources";

        createStartupPresentation(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createStartupPresentation(String resultPath, String resourcesDir) throws Exception {
        // init docbuilder and create new xlsx file
        int doctype = FileTypes.Presentation.PPTX;
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");
        CDocBuilderValue presentation = api.call("GetPresentation");

        // init colors
        CDocBuilderValue backgroundFill = api.call("CreateSolidFill", api.call("CreateRGBColor", 255, 255, 255));
        CDocBuilderValue textFill = api.call("CreateSolidFill", api.call("CreateRGBColor", 80, 80, 80));
        CDocBuilderValue textSpecialFill = api.call("CreateSolidFill", api.call("CreateRGBColor", 15, 102, 7));
        CDocBuilderValue textAltFill = api.call("CreateSolidFill", api.call("CreateRGBColor", 230, 69, 69));
        CDocBuilderValue chartGridFill = api.call("CreateSolidFill", api.call("CreateRGBColor", 134, 134, 134));
        CDocBuilderValue master = presentation.call("GetMaster", 0);
        CDocBuilderValue colorScheme = master.call("GetTheme").call("GetColorScheme");
        colorScheme.call("ChangeColor", 0, api.call("CreateRGBColor", 15, 102, 7));

        // TITLE slide
        CDocBuilderValue slide = presentation.call("GetSlideByIndex", 0);
        slide.call("SetBackground", backgroundFill);
        CDocBuilderValue paragraph = slide.call("GetAllShapes").get(0).call("GetContent").call("GetElement", 0);
        addTextToParagraph(api, paragraph, "GreenVibe Solutions", 120, textSpecialFill, true, "center", "Arial Black");
        paragraph = slide.call("GetAllShapes").get(1).call("GetContent").call("GetElement", 0);
        addTextToParagraph(api, paragraph, "12.12.2024", 48, textFill, false, "center");

        // MARKET OVERVIEW slide
        // parse JSON, obtained as Statista API response
        String jsonPath = resourcesDir + "/data/statista_api_response.json";
        JSONObject data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));
        // create new slide
        slide = addNewSlide(api, backgroundFill);
        // title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(api, paragraph, "Market Overview", 72, textFill, false, "center");
        // market size
        JSONObject market = (JSONObject)data.get("market");
        String[] marketSize = separateValueAndUnit(market.get("size").toString());
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.58);
        addTextToParagraph(api, paragraph, "Market size:", 48, textFill, false, "center");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.97);
        addTextToParagraph(api, paragraph, marketSize[0], 144, textSpecialFill, false, "center", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 3.06);
        addTextToParagraph(api, paragraph, marketSize[1], 48, textFill, false, "center");
        // growth rate
        String[] marketGrowth = separateValueAndUnit(market.get("growth_rate").toString());
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 7, 1.58);
        addTextToParagraph(api, paragraph, "Growth rate:", 48, textFill, false, "center");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 7, 1.97);
        addTextToParagraph(api, paragraph, marketGrowth[0], 144, textSpecialFill, false, "center", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 7, 3.06);
        addTextToParagraph(api, paragraph, marketGrowth[1], 48, textFill, false, "center");
        // trends
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 3.75);
        addTextToParagraph(api, paragraph, "Trends:", 48, textFill, false, "center");
        JSONArray trends = (JSONArray)market.get("trends");
        paragraph = addParagraphToSlide(api, slide, 0.93, 2.92, 1.57, 4.31);
        addTextToParagraph(api, paragraph, makeBulletString('>', trends.size()), 72, textSpecialFill, false, "left", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 9.21, 2.92, 2.1, 4.31);
        String trendsText = "";
        for (Object trend : trends) {
            trendsText += trend.toString() + "\n";
        }
        addTextToParagraph(api, paragraph, trendsText, 72, textSpecialFill, false, "center", "Arial Black");

        // COMPETITORS OVERVIEW section
        // parse JSON, obtained as Statista API response
        jsonPath = resourcesDir + "/data/crunchbase_api_response.json";
        data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));
        // create new slide
        slide = addNewSlide(api, backgroundFill);
        // title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(api, paragraph, "Competitors Overview", 72, textFill, false, "center");
        // chart header
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
        addTextToParagraph(api, paragraph, "Market shares", 48, textFill, false, "center");
        // get chart data
        double othersShare = 100.0;
        JSONArray competitors = (JSONArray)data.get("competitors");
        int competitorsCount = competitors.size();
        CDocBuilderValue shares = context.createArray(competitorsCount + 1);
        CDocBuilderValue competitorNames = context.createArray(competitorsCount + 1);
        for (int i = 0; i < competitorsCount; i++) {
            JSONObject competitor = (JSONObject)competitors.get(i);
            competitorNames.set(i, competitor.get("name").toString());
            String share = competitor.get("market_share").toString();
            // remove last percent symbol
            share = share.substring(0, share.length() - 1);
            othersShare -= Double.parseDouble(share);
            shares.set(i, share);
        }
        shares.set(competitorsCount, Double.toString(othersShare).replace(',', '.'));
        competitorNames.set(competitorsCount, "Others");
        // create a chart
        CDocBuilderValue chartData = context.createArray(1);
        chartData.set(0, shares);
        CDocBuilderValue chart = api.call("CreateChart", "pie", chartData, context.createArray(0), competitorNames);
        setChartSizes(chart, 6.51, 5.9, 4.18, 1.49);
        chart.call("SetLegendFontSize", 14);
        chart.call("SetLegendPos", "right");
        slide.call("AddObject", chart);

        // create slide for every competitor with brief info
        for (int i = 0; i < competitorsCount; i++) {
            JSONObject competitor = (JSONObject)competitors.get(i);
            // create new slide
            slide = addNewSlide(api, backgroundFill);
            // title
            paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            addTextToParagraph(api, paragraph, "Competitors Overview", 72, textFill, false, "center");
            // header
            paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
            addTextToParagraph(api, paragraph, competitor.get("name").toString(), 64, textFill, false, "center");
            // recent funding
            paragraph = addParagraphToSlide(api, slide, 3.13, 0.8, 1.07, 2.65);
            addTextToParagraph(api, paragraph, "Recent funding:", 48, textFill);
            paragraph = addParagraphToSlide(api, slide, 8.9, 0.8, 4.19, 2.52);
            addTextToParagraph(api, paragraph, competitor.get("recent_funding").toString(), 96, textSpecialFill, false, "left", "Arial Black");
            // main products
            paragraph = addParagraphToSlide(api, slide, 3.13, 0.8, 1.07, 3.72);
            addTextToParagraph(api, paragraph, "Main products:", 48, textFill);
            JSONArray products = (JSONArray)competitor.get("products");
            paragraph = addParagraphToSlide(api, slide, 0.93, 3.53, 4.19, 3.72);
            addTextToParagraph(api, paragraph, makeBulletString('>', products.size()), 72, textSpecialFill, false, "left", "Arial Black");
            paragraph = addParagraphToSlide(api, slide, 7.97, 3.53, 5.12, 3.72);
            String productsText = "";
            for (Object product : products) {
                productsText += product.toString() + '\n';
            }
            addTextToParagraph(api, paragraph, productsText, 72, textSpecialFill, false, "left", "Arial Black");
        }

        // TARGET AUDIENCE section
        // parse JSON, obtained as Social Media Insights API response
        jsonPath = resourcesDir + "/data/smi_api_response.json";
        data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));
        // create new slide
        slide = addNewSlide(api, backgroundFill);
        // title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(api, paragraph, "Target Audience", 72, textFill, false, "center");

        // demographics
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.33);
        addTextToParagraph(api, paragraph, "Demographics:", 48, textFill, false, "center");
        // age range
        JSONObject demographics = (JSONObject)data.get("demographics");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.97);
        addTextToParagraph(api, paragraph, demographics.get("age_range").toString(), 128, textSpecialFill, false, "center", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 2.95);
        addTextToParagraph(api, paragraph, "age range", 40, textFill, false, "center");
        // location
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 3.68);
        addTextToParagraph(api, paragraph, demographics.get("location").toString(), 72, textSpecialFill, false, "center", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 4.27);
        addTextToParagraph(api, paragraph, "location", 40, textFill, false, "center");
        // income level
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 5.28);
        addTextToParagraph(api, paragraph, demographics.get("income_level").toString(), 56, textSpecialFill, false, "center", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 5.83);
        addTextToParagraph(api, paragraph, "income level", 40, textFill, false, "center");

        // social trends
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 6, 1.33);
        addTextToParagraph(api, paragraph, "Social trends:", 48, textFill, false, "center");
        // positive feedback
        JSONObject socialTrends = (JSONObject)data.get("social_trends");
        JSONArray positiveFeedbacks = (JSONArray)socialTrends.get("positive_feedback");
        paragraph = addParagraphToSlide(api, slide, 0.63, 2.42, 7, 2.06);
        addTextToParagraph(api, paragraph, makeBulletString('+', positiveFeedbacks.size()), 52, textSpecialFill, false, "left", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.56, 2.42, 7.67, 2.06);
        String positiveFeedback = "";
        for (Object feedback : positiveFeedbacks) {
            positiveFeedback += feedback.toString() + '\n';
        }
        addTextToParagraph(api, paragraph, positiveFeedback, 52, textSpecialFill, false, "left", "Arial Black");
        // negative feedback
        JSONArray negativeFeedbacks = (JSONArray)socialTrends.get("negative_feedback");
        paragraph = addParagraphToSlide(api, slide, 0.63, 2.42, 7, 4.55);
        addTextToParagraph(api, paragraph, makeBulletString('-', negativeFeedbacks.size()), 52, textAltFill, false, "left", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.56, 2.42, 7.67, 4.55);
        String negativeFeedback = "";
        for (Object feedback : negativeFeedbacks) {
            negativeFeedback += feedback.toString() + '\n';
        }
        addTextToParagraph(api, paragraph, negativeFeedback, 52, textAltFill, false, "left", "Arial Black");

        // SEARCH TRENDS section
        // parse JSON, obtained as Google Trends API response
        jsonPath = resourcesDir + "/data/google_trends_api_response.json";
        data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));
        // create new slide
        slide = addNewSlide(api, backgroundFill);
        // title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(api, paragraph, "Search Trends", 72, textFill, false, "center");
        // add every trend on the slide
        JSONArray searchTrends = (JSONArray)data.get("search_trends");
        double offsetY = 1.43;
        for (int i = 0; i < searchTrends.size(); i++) {
            JSONObject searchTrend = (JSONObject)searchTrends.get(i);
            paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, offsetY);
            addTextToParagraph(api, paragraph, searchTrend.get("topic").toString(), 96, textSpecialFill, false, "center", "Arial Black");
            paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, offsetY + 0.8);
            addTextToParagraph(api, paragraph, searchTrend.get("growth").toString(), 40, textFill, false, "center");
            offsetY += 1.25;
        }

        // FINANCIAL MODEL section
        // parse JSON, obtained from financial system
        jsonPath = resourcesDir + "/data/financial_model_data.json";
        data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));
        // create new slide
        slide = addNewSlide(api, backgroundFill);
        // title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(api, paragraph, "Financial Model", 72, textFill, false, "center");
        // chart title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
        addTextToParagraph(api, paragraph, "Profit forecast", 48, textFill, false, "center");
        // chart
        String[] chartKeys = { "revenue", "cost_of_goods_sold", "gross_profit", "operating_expenses", "net_profit" };
        JSONArray profitForecast = (JSONArray)data.get("profit_forecast");
        CDocBuilderValue chartYears = context.createArray(profitForecast.size());
        for (int i = 0; i < profitForecast.size(); i++) {
            JSONObject yearData = (JSONObject)profitForecast.get(i);
            chartYears.set(i, yearData.get("year").toString());
        }
        chartData = context.createArray(chartKeys.length);
        for (int i = 0; i < chartKeys.length; i++) {
            chartData.set(i, context.createArray(profitForecast.size()));
            for (int j = 0; j < profitForecast.size(); j++) {
                JSONObject yearData = (JSONObject)profitForecast.get(j);
                chartData.get(i).set(j, yearData.get(chartKeys[i]).toString());
            }
        }
        CDocBuilderValue chartNames = new CDocBuilderValue(new String[] { "Revenue", "Cost of goods sold", "Gross profit", "Operating expenses", "Net profit" });
        chart = api.call("CreateChart", "lineNormal", chartData, chartNames, chartYears);
        setChartSizes(chart, 10.06, 5.06, 1.67, 2);
        String moneyUnit = separateValueAndUnit(((JSONObject)profitForecast.get(0)).get("revenue").toString())[1];
        String verAxisTitle = "Amount (" + moneyUnit + ")";
        chart.call("SetVerAxisTitle", verAxisTitle, 14, false);
        chart.call("SetHorAxisTitle", "Year", 14, false);
        chart.call("SetLegendFontSize", 14);
        CDocBuilderValue stroke = api.call("CreateStroke", 1, chartGridFill);
        chart.call("SetMinorVerticalGridlines", stroke);
        slide.call("AddObject", chart);

        // break even analysis
        // create new slide
        slide = addNewSlide(api, backgroundFill);
        // title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(api, paragraph, "Financial Model", 72, textFill, false, "center");
        // chart title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
        addTextToParagraph(api, paragraph, "Break even analysis", 48, textFill, false, "center");
        // chart data
        JSONObject breakEvenAnalysis = (JSONObject)data.get("break_even_analysis");
        String[] fixedCostsWithUnit = separateValueAndUnit(breakEvenAnalysis.get("fixed_costs").toString());
        double fixedCosts = Double.parseDouble(fixedCostsWithUnit[0]);
        moneyUnit = fixedCostsWithUnit[1];
        double sellingPricePerUnit = Double.parseDouble(separateValueAndUnit(breakEvenAnalysis.get("selling_price_per_unit").toString())[0]);
        double variableCostPerUnit = Double.parseDouble(separateValueAndUnit(breakEvenAnalysis.get("variable_cost_per_unit").toString())[0]);
        int breakEvenPoint = (int)(long)breakEvenAnalysis.get("break_even_point");
        int step = breakEvenPoint / 4;
        CDocBuilderValue chartUnits = context.createArray(9);
        CDocBuilderValue chartRevenue = context.createArray(9);
        CDocBuilderValue chartTotalCosts = context.createArray(9);
        for (int i = 0; i < 9; i++) {
            int currUnits = i * step;
            chartUnits.set(i, currUnits);
            chartRevenue.set(i, currUnits * sellingPricePerUnit);
            chartTotalCosts.set(i, fixedCosts + currUnits * variableCostPerUnit);
        }
        chartData = new CDocBuilderValue(new CDocBuilderValue[] { chartRevenue, chartTotalCosts });
        chartNames = new CDocBuilderValue(new String[] { "Revenue", "Total costs" });
        // create chart
        chart = api.call("CreateChart", "lineNormal", chartData, chartNames, chartUnits);
        setChartSizes(chart, 9.17, 5.06, 0.31, 2);
        verAxisTitle = "Amount (" + moneyUnit + ")";
        chart.call("SetVerAxisTitle", verAxisTitle, 14, false);
        chart.call("SetHorAxisTitle", "Units sold", 14, false);
        chart.call("SetLegendFontSize", 14);
        chart.call("SetMinorVerticalGridlines", stroke);
        chart.call("SetLegendPos", "bottom");
        slide.call("AddObject", chart);
        // break even point
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 3.11);
        addTextToParagraph(api, paragraph, "Break even point:", 48, textFill, false, "center");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 3.51);
        addTextToParagraph(api, paragraph, Integer.toString(breakEvenPoint), 128, textSpecialFill, false, "center", "Arial Black");
        paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 4.38);
        addTextToParagraph(api, paragraph, "units", 40, textFill, false, "center");

        // growth rates
        // create new slide
        slide = addNewSlide(api, backgroundFill);
        // title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
        addTextToParagraph(api, paragraph, "Financial Model", 72, textFill, false, "center");
        // chart title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
        addTextToParagraph(api, paragraph, "Growth rates", 48, textFill, false, "center");
        // chart
        JSONArray growthRates = (JSONArray)data.get("growth_rates");
        chartYears = context.createArray(growthRates.size());
        CDocBuilderValue chartGrowth = context.createArray(growthRates.size());
        for (int i = 0; i < growthRates.size(); i++) {
            JSONObject yearData = (JSONObject)growthRates.get(i);
            chartYears.set(i, yearData.get("year").toString());
            chartGrowth.set(i, yearData.get("growth").toString());
        }
        chartData = new CDocBuilderValue( new CDocBuilderValue[] { chartGrowth });
        chart = api.call("CreateChart", "lineNormal", chartData, context.createArray(0), chartYears);
        setChartSizes(chart, 10.06, 5.06, 1.67, 2);
        chart.call("SetVerAxisTitle", "Growth (%)", 14, false);
        chart.call("SetHorAxisTitle", "Year", 14, false);
        chart.call("SetMinorVerticalGridlines", stroke);
        slide.call("AddObject", chart);

        // save and close
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }

    public static void addTextToParagraph(CDocBuilderValue api, CDocBuilderValue paragraph, String text, int fontSize, CDocBuilderValue fill, boolean isBold, String jc, String fontFamily) {
        CDocBuilderValue run = api.call("CreateRun");
        run.call("AddText", text);
        run.call("SetFontSize", fontSize);
        run.call("SetBold", isBold);
        run.call("SetFill", fill);
        run.call("SetFontFamily", fontFamily);
        paragraph.call("AddElement", run);
        paragraph.call("SetJc", jc);
    }

    public static void addTextToParagraph(CDocBuilderValue api, CDocBuilderValue paragraph, String text, int fontSize, CDocBuilderValue fill, boolean isBold, String jc) {
        addTextToParagraph(api, paragraph, text, fontSize, fill, isBold, jc, "Arial");
    }

    public static void addTextToParagraph(CDocBuilderValue api, CDocBuilderValue paragraph, String text, int fontSize, CDocBuilderValue fill, boolean isBold) {
        addTextToParagraph(api, paragraph, text, fontSize, fill, isBold, "left", "Arial");
    }

    public static void addTextToParagraph(CDocBuilderValue api, CDocBuilderValue paragraph, String text, int fontSize, CDocBuilderValue fill) {
        addTextToParagraph(api, paragraph, text, fontSize, fill, false, "left", "Arial");
    }

    public static CDocBuilderValue addNewSlide(CDocBuilderValue api, CDocBuilderValue fill) {
        CDocBuilderValue slide = api.call("CreateSlide");
        CDocBuilderValue presentation = api.call("GetPresentation");
        presentation.call("AddSlide", slide);
        slide.call("SetBackground", fill);
        slide.call("RemoveAllObjects");
        return slide;
    }

    public static final int em_in_inch = 914400;
    // width, height, pos_x and pos_y are set in INCHES
    public static CDocBuilderValue addParagraphToSlide(CDocBuilderValue api, CDocBuilderValue slide, double width, double height, double pos_x, double pos_y) {
        CDocBuilderValue shape = api.call("CreateShape", "rect", width * em_in_inch, height * em_in_inch);
        shape.call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
        CDocBuilderValue paragraph = shape.call("GetDocContent").call("GetElement", 0);
        slide.call("AddObject", shape);
        return paragraph;
    }

    public static void setChartSizes(CDocBuilderValue chart, double width, double height, double pos_x, double pos_y) {
        chart.call("SetSize", width * em_in_inch, height * em_in_inch);
        chart.call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
    }

    public static String[] separateValueAndUnit(String data) {
        int spacePos = data.indexOf(' ');
        return new String[] { data.substring(0, spacePos), data.substring(spacePos + 1) };
    }

    public static String makeBulletString(char bullet, int repeats) {
        String result = "";
        for (int i = 0; i < repeats; i++) {
            result += bullet;
            result += '\n';
        }
        return result;
    }
}
