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

import os
import sys
sys.path.append('../../out/python')
import constants
sys.path.append(constants.BUILDER_DIR)
import docbuilder

def createImageSlide(api, presentation, image_url):
    slide = api.Call('CreateSlide')
    presentation.Call('AddSlide', slide)
    fill = api.Call('CreateBlipFill', image_url, 'stretch')
    slide.Call('SetBackground', fill)
    slide.Call('RemoveAllObjects')
    return slide

def addTextToSlideShape(api, content, text, fontSize, isBold, js):
    paragraph = api.Call('CreateParagraph')
    paragraph.Call('SetSpacingBefore', 0)
    paragraph.Call('SetSpacingAfter', 0)
    content.Call('Push', paragraph)
    run = paragraph.Call('AddText', text)
    run.Call('SetFill', api.Call('CreateSolidFill', api.Call('CreateRGBColor', 0xff, 0xff, 0xff)))
    run.Call('SetFontSize', fontSize)
    run.Call('SetFontFamily', 'Georgia')
    run.Call('SetBold', isBold)
    paragraph.Call('SetJc', js)

if __name__ == '__main__':
    slideImages = {}
    slideImages['gun'] = 'https://static.onlyoffice.com/assets/docs/samples/img/presentation_gun.png'
    slideImages['axe'] = 'https://static.onlyoffice.com/assets/docs/samples/img/presentation_axe.png'
    slideImages['knight'] = 'https://static.onlyoffice.com/assets/docs/samples/img/presentation_knight.png'
    slideImages['sky'] = 'https://static.onlyoffice.com/assets/docs/samples/img/presentation_sky.png'

    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Presentation.PPTX)

    context = builder.GetContext()
    globalObj = context.GetGlobal()
    api = globalObj['Api']

    # Create presentation
    presentation = api.Call('GetPresentation')
    presentation.Call('SetSizes', 9144000, 6858000)

    slide = createImageSlide(api, presentation, slideImages['gun'])
    presentation.Call('GetSlideByIndex', 0).Call('Delete')

    shape = api.Call('CreateShape', 'rect', 8056800, 3020400, api.Call('CreateNoFill'), api.Call('CreateStroke', 0, api.Call('CreateNoFill')))
    shape.Call('SetPosition', 608400, 1267200)
    content = shape.Call('GetDocContent')
    content.Call('RemoveAllElements')
    addTextToSlideShape(api, content, 'How They', 160, True, 'left')
    addTextToSlideShape(api, content, 'Throw Out', 132, False, 'left')
    addTextToSlideShape(api, content, 'a Challenge', 132, False, 'left')
    slide.Call('AddObject', shape)

    slide = createImageSlide(api, presentation, slideImages['axe'])

    shape = api.Call('CreateShape', 'rect', 6904800, 1724400, api.Call('CreateNoFill'), api.Call('CreateStroke', 0, api.Call('CreateNoFill')))
    shape.Call('SetPosition', 1764000, 1191600)
    content = shape.Call('GetDocContent')
    content.Call('RemoveAllElements')
    addTextToSlideShape(api, content, 'American Indians ', 110, True, 'right')
    addTextToSlideShape(api, content, '(XVII century)', 94, False, 'right')
    slide.Call('AddObject', shape)

    shape = api.Call('CreateShape', 'rect', 4986000, 2419200, api.Call('CreateNoFill'), api.Call('CreateStroke', 0, api.Call('CreateNoFill')))
    shape.Call('SetPosition', 3834000, 3888000)
    content = shape.Call('GetDocContent')
    content.Call('RemoveAllElements')
    addTextToSlideShape(api, content, 'put a tomahawk on the ground in the ', 84, False, 'right')
    addTextToSlideShape(api, content, 'rival\'s camp', 84, False, 'right')
    slide.Call('AddObject', shape)

    slide = createImageSlide(api, presentation, slideImages['knight'])

    shape = api.Call('CreateShape', 'rect', 6904800, 1724400, api.Call('CreateNoFill'), api.Call('CreateStroke', 0, api.Call('CreateNoFill')))
    shape.Call('SetPosition', 1764000, 1191600)
    content = shape.Call('GetDocContent')
    content.Call('RemoveAllElements')
    addTextToSlideShape(api, content, 'European Knights', 110, True, 'right')
    addTextToSlideShape(api, content, ' (XII-XVI centuries)', 94, False, 'right')
    slide.Call('AddObject', shape)

    shape = api.Call('CreateShape', 'rect', 4986000, 2419200, api.Call('CreateNoFill'), api.Call('CreateStroke', 0, api.Call('CreateNoFill')))
    shape.Call('SetPosition', 3834000, 3888000)
    content = shape.Call('GetDocContent')
    content.Call('RemoveAllElements')
    addTextToSlideShape(api, content, 'threw a glove', 84, False, 'right')
    addTextToSlideShape(api, content, 'in the rival\'s face', 84, False, 'right')
    slide.Call('AddObject', shape)

    slide = createImageSlide(api, presentation, slideImages['sky'])

    shape = api.Call('CreateShape', 'rect', 7887600, 3063600, api.Call('CreateNoFill'), api.Call('CreateStroke', 0, api.Call('CreateNoFill')))
    shape.Call('SetPosition', 630000, 1357200)
    content = shape.Call('GetDocContent')
    content.Call('RemoveAllElements')
    addTextToSlideShape(api, content, 'OnlyOffice', 176, False, 'center')
    addTextToSlideShape(api, content, 'stands for Peace', 132, False, 'center')
    slide.Call('AddObject', shape)

    # Save and close
    resultPath = os.getcwd() + '/result.pptx'
    builder.SaveFile(docbuilder.FileTypes.Presentation.PPTX, resultPath)
    builder.CloseFile()
