'''
(c) Copyright Ascensio System SIA 2024

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
'''

# import docbuilder
import sys
from datetime import datetime

sys.path.append('../../out/python')
import constants
sys.path.append(constants.BUILDER_DIR)
import docbuilder
# import other standard libraries
import json
import os


def set_table_style(range):
    range.Call('SetRowHeight', 24)
    range.Call('SetAlignVertical', 'center')

    color = api.Call('CreateColorFromRGB', 0, 0, 0)
    line_style = 'Thin'
    range.Call('SetBorders', 'Top', line_style, color)
    range.Call('SetBorders', 'Left', line_style, color)
    range.Call('SetBorders', 'Right', line_style, color)
    range.Call('SetBorders', 'Bottom', line_style, color)
    range.Call('SetBorders', 'InsideHorizontal', line_style, color)
    range.Call('SetBorders', 'InsideVertical', line_style, color)

def get_average_values(feedback_data):
    result = {}
    for item in feedback_data:
        for key, value in item['answers'].items():
            if key not in result:
                result[key] = [value['rating']]
            else:
                result[key].append(value['rating'])

    average = [['Question', 'Average Rating', 'Number of Responses']]
    for key, value in result.items():
        average.append([key, round(sum(value) / len(value), 1), len(value)])

    return average


def fill_average_sheet(worksheet, feedback_data):
    average_values = get_average_values(feedback_data)
    cols_count = len(average_values[0]) - 1
    start_sell = worksheet.Call('GetRangeByNumber', 0, 0)
    end_cell = worksheet.Call('GetRangeByNumber', len(average_values) - 1, cols_count)

    average_range = worksheet.Call('GetRange', start_sell, end_cell)
    set_table_style(average_range)
    worksheet.Call('GetRange', worksheet.Call('GetRangeByNumber', 1, 1), end_cell).Call('SetAlignHorizontal', 'center')

    header_row = worksheet.Call('GetRange', start_sell, worksheet.Call('GetRangeByNumber', 0, cols_count))
    header_row.Call('SetBold', True)

    average_range.Call('SetValue', average_values)
    average_range.Call('AutoFit', False, True)

    return len(average_values)


def fill_personal_ratings_and_comments(worksheet, feedback_data):
    header_values = [['Date', 'Question', 'Comment', 'Rating', 'Average User Rating']]
    cols_count = len(header_values[0]) - 1
    start_sell = worksheet.Call('GetRangeByNumber', 0, 0)

    header_row = worksheet.Call('GetRange', start_sell, worksheet.Call('GetRangeByNumber', 0, cols_count))
    header_row.Call('SetValue', header_values)
    header_row.Call('SetBold', True)

    rows_count = 1
    for item in feedback_data:
        # Count and fill user feedback
        user_feedback = []
        rating = 0
        for key, value in item['answers'].items():
            user_feedback.append([key, value['comment'], value['rating']])
            rating += value['rating']
        
        user_rows_count = len(user_feedback) - 1
        # Fill date
        rating_cell = worksheet.Call(
            'GetRange',
            worksheet.Call('GetRangeByNumber', rows_count, 0),
            worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, 0),
        )
        rating_cell.Call('Merge', False)
        rating_cell.Call('SetValue', item['date'])

        # Fill ratings
        user_range = worksheet.Call(
            'GetRange',
            worksheet.Call('GetRangeByNumber', rows_count, 1),
            worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, cols_count - 1),
        )
        user_range.Call('SetValue', user_feedback)

        # Count average rating
        rating = round(rating/len(user_feedback), 1)
        rating_cell = worksheet.Call(
            'GetRange',
            worksheet.Call('GetRangeByNumber', rows_count, cols_count),
            worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, cols_count),
        )
        rating_cell.Call('Merge', False)
        rating_cell.Call('SetValue', rating)

        # If rating <= 2, highlight it
        if rating <= 2:
            worksheet.Call(
                'GetRange',
                worksheet.Call('GetRangeByNumber', rows_count, 0),
                worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, cols_count),
            ).Call('SetFillColor', api.Call('CreateColorFromRGB', 237, 125, 49))

        # Update rows count
        rows_count += len(user_feedback)

    # Format table
    rows_count -= 1
    result_range = worksheet.Call(
        'GetRange',
        start_sell,
        worksheet.Call('GetRangeByNumber', rows_count, cols_count),
    )
    set_table_style(result_range)
    worksheet.Call(
        'GetRange',
        worksheet.Call('GetRangeByNumber', 1, cols_count - 1),
        worksheet.Call('GetRangeByNumber', rows_count, cols_count),
    ).Call('SetAlignHorizontal', 'center')
    result_range.Call('AutoFit', False, True)

    return rows_count + 1


def create_column_chart(worksheet, data_range, title):
    chart = worksheet.Call('AddChart', data_range, False, 'bar', 2, 135.38 * 36000, 81.28 * 36000)
    chart.Call('SetPosition', 0, 0, 0, 0)
    chart.Call('SetTitle', title, 16)


def create_line_chart(api, worksheet, data_range, title):
    chart = worksheet.Call('AddChart', data_range, False, 'scatter', 2, 135.38 * 36000, 81.28 * 36000)
    chart.Call('SetPosition', 0, 0, 18, 0)
    stroke = api.Call(
        'CreateStroke',
        0.5 * 36000,
        api.Call('CreateSolidFill', api.Call('CreateRGBColor', 128, 128, 128)),
    )
    chart.Call('SetSeriesOutLine', stroke, 0, False)
    chart.Call('SetTitle', title, 16)
    chart.Call('SetMajorHorizontalGridlines', api.Call('CreateStroke', 0,  api.Call('CreateNoFill')))


def create_pie_chart(api, worksheet, data_range, title):
    worksheet.Call('GetRange', '$A$1:$C$2').Call(
        'SetValue',
        [
            ['Negative', 'Neutral', 'Positive'],
            [f'=COUNTIF({data_range}, "<=2")', f'=COUNTIF({data_range}, "=3")', f'=COUNTIF({data_range}, ">=4")'],
        ],
    )
    chart = worksheet.Call('AddChart', 'Charts!$A$1:$C$2', True, 'pie', 2, 135.38 * 36000, 81.28 * 36000)
    chart.Call('SetPosition', 9, 0, 0, 0)
    chart.Call('SetTitle', title, 16)
    chart.Call('SetDataPointFill', api.Call('CreateSolidFill', api.Call('CreateRGBColor', 237, 125, 49)), 0, 0)
    chart.Call('SetDataPointFill', api.Call('CreateSolidFill', api.Call('CreateRGBColor', 128, 128, 128)), 0, 1)
    chart.Call('SetDataPointFill', api.Call('CreateSolidFill', api.Call('CreateRGBColor', 91, 155, 213)), 0, 2)
    stroke = api.Call(
        'CreateStroke',
        0.5 * 36000,
        api.Call('CreateSolidFill', api.Call('CreateRGBColor', 255, 255, 255)),
    )
    chart.Call('SetSeriesOutLine', stroke, 0, False)


if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/user_feedback_report_response.json'), 'r') as file_json:
        data = json.load(file_json)

    # Sort feedback data
    data.sort(key=lambda k: datetime.strptime(k['date'], '%Y-%m-%d'))

    # init docbuilder and create new pdf file
    doctype = docbuilder.FileTypes.Spreadsheet.XLSX
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(doctype)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']

    # Get current worksheet
    worksheet1 = api.Call('GetActiveSheet')

    # Create worksheet with average values
    worksheet1.Call('SetName', 'Average')
    table1_rows_count = fill_average_sheet(worksheet1, data)

    # Create worksheet with comments and personal ratings
    api.Call('AddSheet', 'Comments')
    worksheet2 = api.Call('GetActiveSheet')
    table2_rows_count = fill_personal_ratings_and_comments(worksheet2, data)

    # Create worksheet with charts
    api.Call('AddSheet', 'Charts')
    worksheet3 = api.Call('GetActiveSheet')
    create_column_chart(worksheet3, f'Average!$A$2:$B${table1_rows_count}', 'Average ratings')
    create_line_chart(
        api,
        worksheet3,
        f'Comments!$A$1:$A${table2_rows_count},Comments!$E$1:$E${table2_rows_count}',
        'Dynamics of the average ratings',
    )
    create_pie_chart(api, worksheet3, f'Comments!$D$1:$D${table2_rows_count}', 'Shares of reviews')

    # Set first worksheet active
    worksheet1.Call('SetActive')

    # save and close
    result_path = os.getcwd() + '/result.xlsx'
    builder.SaveFile(doctype, result_path)
    builder.CloseFile()
