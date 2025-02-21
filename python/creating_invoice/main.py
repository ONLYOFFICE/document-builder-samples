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
sys.path.append('../../out/python')
import constants
sys.path.append(constants.BUILDER_DIR)
import docbuilder
# import other standard libraries
import json
import os


def fill_header(paragraph, text):
    paragraph.Call('AddText', text)
    paragraph.Call('SetFontSize', 28)
    paragraph.Call('SetBold', True)


def set_spacing_after(paragraph, spacing):
    paragraph.Call('SetSpacingAfter', spacing)


def set_numbering(paragraph, num_lvl):
    paragraph.Call('SetNumbering', num_lvl)


def create_details_header(api, text):
    paragraph = api.Call('CreateParagraph')
    paragraph.Call('AddText', text)
    paragraph.Call('SetBold', True)
    paragraph.Call('SetItalic', True)
    set_spacing_after(paragraph, 50)
    return paragraph


def setup_requisites_style(paragraph, num_lvl, set_spacing):
    if set_spacing:
        set_spacing_after(paragraph, 20)
    if num_lvl:
        set_numbering(paragraph, num_lvl)


def create_requisites_paragraph(api, title, details, num_lvl=None, set_spacing=True, set_title_bold=True):
    paragraph = api.Call('CreateParagraph')
    title_run = paragraph.Call('AddText', f'{title}: ')
    if set_title_bold:
        title_run.Call('SetBold', True)
    else:
        title_run.Call('SetItalic', True)
    details_run = paragraph.Call('AddText', details)
    details_run.Call('SetItalic', True)
    setup_requisites_style(paragraph, num_lvl, set_spacing)
    return paragraph


def setup_table_style(document, table):
    # table size
    table.Call('SetWidth', 'percent', 100)
    table.Call('Select')
    table_range = document.Call('GetRangeBySelect')
    table_paragraphs = table_range.Call('GetAllParagraphs')
    for i in range(table_paragraphs.GetLength()):
        para_pr = table_paragraphs.Get(i).Call('GetParaPr')
        para_pr.Call('SetSpacingBefore', 40)
        para_pr.Call('SetSpacingAfter', 40)

    # table borders
    params = ("single", 4, 0, 0, 0, 0)
    table.Call('SetTableBorderTop', *params)
    table.Call('SetTableBorderBottom', *params)
    table.Call('SetTableBorderLeft', *params)
    table.Call('SetTableBorderRight', *params)
    table.Call('SetTableBorderInsideH', *params)
    table.Call('SetTableBorderInsideV', *params)


def get_cell_content(cell):
    return cell.Call('GetContent').Call('GetElement', 0)


def fill_table_content(table, items):
    table_headers = ['Description', 'Quantity', 'Unit Price', 'Total']
    table_fields = ['description', 'quantity', 'unit_price', 'total']

    # fill table header
    header_row = table.Call('GetRow', 0)
    for i, field_name in enumerate(table_headers):
        header_cell = get_cell_content(header_row.Call('GetCell', i))
        header_cell.Call('AddText', field_name)
        header_cell.Call('SetBold', True)

    # fill items
    items.append({key: '...' for key in table_fields})
    for i in range(len(items)):
        row = table.Call('GetRow', i + 1)
        for j, key in enumerate(table_fields):
            cell = get_cell_content(row.Call('GetCell', j))
            cell.Call('AddText', str(items[i][key]))


if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/invoice_response.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new pdf file
    doctype = docbuilder.FileTypes.Document.OFORM_PDF
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(doctype)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    document = api.Call('GetDocument')

    # DOCUMENT STYLE
    text_pr = document.Call('GetDefaultTextPr')
    text_pr.Call('SetFontSize', 24)
    text_pr.Call('SetFontFamily', 'Times New Roman')

    # DOCUMENT HEADER
    header = document.Call('GetElement', 0)
    fill_header(header, 'INVOICE')

    # document requisites
    document.Call('Push', create_requisites_paragraph(api, 'Invoice No.', data["invoice"]["number"]))
    document.Call('Push', create_requisites_paragraph(api, 'Date', data["invoice"]["date"], set_spacing=False))

    # bullet numbering
    bullet_numbering = document.Call('CreateNumbering', 'bullet')
    num_lvl_1 = bullet_numbering.Call('GetLevel', 0)

    # SELLER INFORMATION
    seller_header = create_details_header(api, 'SELLER INFORMATION')
    document.Call('Push', seller_header)

    # seller details
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Company Name', data["seller"]["company_name"], num_lvl=num_lvl_1),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Address', data["seller"]["address"], num_lvl=num_lvl_1),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Tax ID (TIN)', data["seller"]["tin"], num_lvl=num_lvl_1),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Bank Details', '', num_lvl=num_lvl_1),
    )

    # bank details
    num_lvl_2 = bullet_numbering.Call('GetLevel', 1)
    num_lvl_2.Call('SetCustomType', 'none', '', 'left')
    num_lvl_2.Call('SetSuff', 'space')

    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Bank Name',
            data["seller"]["bank_details"]["bank_name"],
            num_lvl=num_lvl_2,
            set_title_bold=False,
        )
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Account Number',
            data["seller"]["bank_details"]["account_number"],
            num_lvl=num_lvl_2,
            set_title_bold=False,
        )
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'SWIFT Code',
            data["seller"]["bank_details"]["swift_code"],
            num_lvl=num_lvl_2,
            set_spacing=False,
            set_title_bold=False,
        )
    )

    # BUYER INFORMATION
    buyer_header = create_details_header(api, 'BUYER INFORMATION')
    document.Call('Push', buyer_header)

    # buyer details
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Company Name', data["buyer"]["company_name"], num_lvl=num_lvl_1),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Address', data["buyer"]["address"], num_lvl=num_lvl_1),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Tax ID (TIN)',
            data["buyer"]["tin"],
            num_lvl=num_lvl_1,
            set_spacing=False,
        )
    )

    # TABLE OF ITEMS
    table_header = api.Call('CreateParagraph')
    fill_header(table_header, 'TABLE OF ITEMS')
    document.Call('Push', table_header)

    # table content
    items_table = api.Call('CreateTable', 4, len(data["items"]) + 2)
    document.Call('Push', items_table)
    setup_table_style(document, items_table)
    fill_table_content(items_table, data["items"])

    # TOTALS
    totals = create_details_header(api, 'TOTALS')
    document.Call('Push', totals)
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Subtotal', f'${data["totals"]["subtotal"]}', num_lvl=num_lvl_1),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Tax (20% VAT)', f'${data["totals"]["tax"]}', num_lvl=num_lvl_1),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Total Amount Due',
            f'${data["totals"]["total_due"]}',
            num_lvl=num_lvl_1,
            set_spacing=False,
        )
    )

    # SIGNATURE
    sign_header = api.Call('CreateParagraph')
    sign_header.Call('AddText', 'Signature:')
    sign_header.Call('SetBold', True)
    document.Call('Push', sign_header)

    sign_details = api.Call('CreateParagraph')
    sign_details.Call('AddText', f'{data["seller"]["authorized_person"]}, {data["seller"]["position"]}')
    sign_details.Call('AddLineBreak')
    sign_details.Call('AddText', data["seller"]["company_name"])
    document.Call('Push', sign_details)

    # save and close
    result_path = os.getcwd() + '/result.pdf'
    builder.SaveFile(doctype, result_path)
    builder.CloseFile()
