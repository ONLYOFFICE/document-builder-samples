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

# Helper functions
def setPictureFormProperties(pictureForm, key, tip, required, placeholder, scaleFlag, lockAspectRatio, respectBorders, shiftX, shiftY):
    pictureForm.Call('SetFormKey', key)
    pictureForm.Call('SetTipText', tip)
    pictureForm.Call('SetRequired', required)
    pictureForm.Call('SetPlaceholderText', placeholder)
    pictureForm.Call('SetScaleFlag', scaleFlag)
    pictureForm.Call('SetLockAspectRatio', lockAspectRatio)
    pictureForm.Call('SetRespectBorders', respectBorders)
    pictureForm.Call('SetPicturePosition', shiftX, shiftY)

def setTextFormProperties(textForm, key, tip, required, placeholder, comb, maxCharacters, cellWidth, multiLine, autoFit):
    textForm.Call('SetFormKey', key)
    textForm.Call('SetTipText', tip)
    textForm.Call('SetRequired', required)
    textForm.Call('SetPlaceholderText', placeholder)
    textForm.Call('SetComb', comb)
    textForm.Call('SetCharactersLimit', maxCharacters)
    textForm.Call('SetCellWidth', cellWidth)
    textForm.Call('SetCellWidth', multiLine)
    textForm.Call('SetMultiline', autoFit)

if __name__ == '__main__':
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Document.DOCX)

    context = builder.GetContext()
    globalObj = context.GetGlobal()
    api = globalObj['Api']

    document = api.Call('GetDocument')
    paragraph = document.Call('GetElement', 0)
    headingStyle = document.Call('GetStyle', 'Heading 3')

    paragraph.Call('AddText', 'Employee pass card')
    paragraph.Call('SetStyle', headingStyle)
    document.Call('Push', paragraph)

    pictureForm = api.Call('CreatePictureForm')
    setPictureFormProperties(pictureForm, 'Photo', 'Upload your photo', False, 'Photo', 'tooBig', True, False, 50, 50)
    paragraph = api.Call('CreateParagraph')
    paragraph.Call('AddElement', pictureForm)
    document.Call('Push', paragraph)

    textForm = api.Call('CreateTextForm')
    setTextFormProperties(textForm, 'First name', 'Enter your first name', False, 'First name', True, 13, 3, False, False)
    paragraph = api.Call('CreateParagraph')
    paragraph.Call('AddElement', textForm)
    document.Call('Push', paragraph)

    # Save and close
    resultPath = os.getcwd() + '/result.docx'
    builder.SaveFile(docbuilder.FileTypes.Document.DOCX, resultPath)
    builder.CloseFile()
