# converts multiple word docs to pdfs

inpath = 'L:\\Research\\Solar Research\\SolarStateProfiles\\New WhatYouNeedToKnow\\'
outpath = 'L:\\Research\\Solar Research\\SolarStateProfiles\\New WhatYouNeedToKnow\\'

import sys
import os
import comtypes.client

wdFormatPDF = 17

# %% Get files in input folder
input_file_paths = os.listdir(inpath)

#%% Convert each file
for input_file_name in input_file_paths:

    # Skip if file does not contain a power point extension
    if not input_file_name.lower().endswith((".docx")):
        continue

    # Create input file path
    input_file_path = os.path.join(inpath, input_file_name)
    print(input_file_path)

    # Create powerpoint application object
    word = comtypes.client.CreateObject('Word.Application')

    # Set visibility to minimize
    word.Visible = 1

    # Open the powerpoint slides
    doc = word.Documents.Open(input_file_path)

    # Get base file name
    file_name = os.path.splitext(input_file_name)[0]

    # Create output file path
    output_file_path = outpath + file_name + ".pdf"
    print(output_file_path)

    # Save as PDF (formatType = 32)
    doc.SaveAs(output_file_path, FileFormat=wdFormatPDF)

    # Close the slide deck
    doc.Close()




#in_file = os.path.abspath(sys.argv[1])
#out_file = os.path.abspath(sys.argv[2])

#word = comtypes.client.CreateObject('Word.Application')
#doc = word.Documents.Open(in_file)
#doc.SaveAs(out_file, FileFormat=wdFormatPDF)
#doc.Close()
#word.Quit()
