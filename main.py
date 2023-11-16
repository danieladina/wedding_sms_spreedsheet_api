
import os
from daniel import Create_Service

from sheet_design import sheet_body
import time



start_time = time.time()

FOLDER_PATH="C:\\Users\\danie\\Desktop\\sms_project\\create"
CLIENT_SECRET_FILE=os.path.join(FOLDER_PATH, 'credentials.json')
API_SERVICE_NAME='sheets'
API_VERSION='v4'
SCOPES=['https://www.googleapis.com/auth/spreadsheets']

# Create the service
service=Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

sheetfile2=service.spreadsheets().create(body=sheet_body).execute()
print(sheetfile2['spreadsheetUrl'])
print(sheetfile2['spreadsheetId'])
SHEET_NAME='החתונה של פלוני ואלמונית'
for _sheet in sheetfile2['sheets']:
    sheet_id=_sheet['properties']['sheetId']

Id=sheetfile2['spreadsheetId']
num_id=sheet_id



border_style = {
    "style": "SOLID",  # You can use other styles like DOTTED, DASHED, etc.
    "width": 1,        # Border width (1 is the default)
    "color": {"red": 0.0, "green": 0.0, "blue": 0.0, "alpha": 1.0}  # Black color
}




# Merge cells for the first 3 rows and 6 columns
values=[
    ['מעקב אחר הזמנות לרשימת אורחים', ]
]
work_sheet_name='sheet1!'
cell_range_insert='A1'

merge_range={
    "sheetId": num_id,
    "startRowIndex": 0,
    "endRowIndex": 3,
    "startColumnIndex": 0,
    "endColumnIndex": 4
}

# Create a merge request
merge_request={
    "mergeCells": {
        "range": merge_range,
        "mergeType": "MERGE_ALL"
    }
}

# Define the cell range to merge (E1:F1)
merge_range_e1_f1={
    "sheetId": num_id,
    "startRowIndex": 0,  # Row index for E1
    "endRowIndex": 1,  # Row index for F1
    "startColumnIndex": 4,  # Column E
    "endColumnIndex": 6,  # Column F
}

# Create a merge request for E1:F1
merge_request_e1_f1={
    "mergeCells": {
        "range": merge_range_e1_f1,
        "mergeType": "MERGE_ALL"
    }
}

# Merge cells E2, E3, F2, F3
merge_range_e2_f3={
    "sheetId": num_id,
    "startRowIndex": 1,  # Row index for E2
    "endRowIndex": 3,  # End row index for E3
    "startColumnIndex": 4,  # Column E
    "endColumnIndex": 6  # Column F
}

merge_request_e2_f3={
    "mergeCells": {
        "range": merge_range_e2_f3,
        "mergeType": "MERGE_ALL"
    }
}


# Define the cell range to merge (g2:g3)
merge_range_g2_g3={
    "sheetId": num_id,
    "startRowIndex": 1,  # Row index for E1
    "endRowIndex": 3,  # Row index for F1
    "startColumnIndex": 6,  # Column E
    "endColumnIndex": 7,  # Column F
}

# Create a merge request for G2:G3
merge_request_g2_g3={
    "mergeCells": {
        "range": merge_range_g2_g3,
        "mergeType": "MERGE_ALL"
    }
}


# Define the range to merge (H2:H3)
merge_range_h2_h3 = {
    "sheetId": num_id,
    "startRowIndex": 1,  # Row index for H2
    "endRowIndex": 3,  # Row index for H3
    "startColumnIndex": 7,  # Column H
    "endColumnIndex": 8  # Column H
}

# Create a merge request for H2:H3
merge_request_h2_h3 = {
    "mergeCells": {
        "range": merge_range_h2_h3,
        "mergeType": "MERGE_ALL"
    }
}


# Define the merge range for J2:J4
merge_range_j2_j4 = {
    "sheetId": num_id,
    "startRowIndex": 1,    # Row index for J2
    "endRowIndex": 4,      # Row index for J4 (inclusive)
    "startColumnIndex": 9, # Column J
    "endColumnIndex": 10   # Column J
}

# Create a merge request for J2:J4
merge_request_j2_j4 = {
    "mergeCells": {
        "range": merge_range_j2_j4,
        "mergeType": "MERGE_ALL"
    }
}





# Define the cell range to merge (g2:g3)
merge_range_i2_i1000={
    "sheetId": num_id,
    "startRowIndex": 0,  # Row index for E1
    "endRowIndex": 1000,  # Row index for F1
    "startColumnIndex": 8,  # Column E
    "endColumnIndex": 9,  # Column F
}

# Create a merge request for G2:G3
merge_request_i2_i10000={
    "mergeCells": {
        "range": merge_range_i2_i1000,
        "mergeType": "MERGE_ALL"
    }
}
# Define the cell range to merge (g2:g3)
merge_range_k2_k1000={
    "sheetId": num_id,
    "startRowIndex": 0,  # Row index for E1
    "endRowIndex": 1000,  # Row index for F1
    "startColumnIndex": 10,  # Column E
    "endColumnIndex": 11,  # Column F
}

# Create a merge request for G2:G3
merge_request_k2_k10000={
    "mergeCells": {
        "range": merge_range_k2_k1000,
        "mergeType": "MERGE_ALL"
    }
}



# Define the cell range to merge (j7:j8)
merge_range_j7_j8 = {
    "sheetId": num_id,
    "startRowIndex": 6,  # Row index for J7
    "endRowIndex": 9,    # Row index for J8
    "startColumnIndex": 9,  # Column J
    "endColumnIndex": 10,
}

# Create a merge request for J7:J8
merge_request_j7_j8 = {
    "mergeCells": {
        "range": merge_range_j7_j8,
        "mergeType": "MERGE_ALL"
    }
}




# Define the cell range to merge (J12:J14)
merge_range_j12_j14 = {
    "sheetId": num_id,
    "startRowIndex": 11,  # Row index for J11
    "endRowIndex": 14,    # Row index for J13
    "startColumnIndex": 9,  # Column J
    "endColumnIndex": 10,
}

# Create a merge request for J11:J13
merge_request_j12_j14 = {
    "mergeCells": {
        "range": merge_range_j12_j14,
        "mergeType": "MERGE_ALL"
    }
}

# Define the cell range to merge (J17:J19)
merge_range_j17_j19 = {
    "sheetId": num_id,
    "startRowIndex": 16,  # Row index for J17
    "endRowIndex": 19,    # Row index for J19
    "startColumnIndex": 9,  # Column J
    "endColumnIndex": 10,
}

# Create a merge request for J17:J19
merge_request_j17_j19 = {
    "mergeCells": {
        "range": merge_range_j17_j19,
        "mergeType": "MERGE_ALL"
    }
}



# Set the text direction to right-to-left (RTL)
text_direction_request={
    "updateSheetProperties": {
        "properties": {
            "sheetId": num_id,
            "rightToLeft": True
        },
        "fields": "rightToLeft"
    }
}

# Resize the i column to a specific width (e.g., 10 pixels)
resize_column_request_i_column = {
    "updateDimensionProperties": {
        "range": {
            "sheetId": num_id,  # Replace with the correct sheetId
            "dimension": "COLUMNS",
            "startIndex": 8,  # 0-based index of the F column
            "endIndex": 9
        },
        "properties": {
            "pixelSize": 10  # Adjust the pixelSize as needed
        },
        "fields": "pixelSize"
    }
}
# Resize the j column to a specific width (e.g., 10 pixels)
resize_column_request_j_column = {
    "updateDimensionProperties": {
        "range": {
            "sheetId": num_id,  # Replace with the correct sheetId
            "dimension": "COLUMNS",
            "startIndex": 9,  # 0-based index of the F column
            "endIndex": 10
        },
        "properties": {
            "pixelSize": 150 # Adjust the pixelSize as needed
        },
        "fields": "pixelSize"
    }
}
# Resize the k column to a specific width (e.g., 10 pixels)
resize_column_request_k_column = {
    "updateDimensionProperties": {
        "range": {
            "sheetId": num_id,  # Replace with the correct sheetId
            "dimension": "COLUMNS",
            "startIndex": 10,  # 0-based index of the F column
            "endIndex": 11
        },
        "properties": {
            "pixelSize": 30 # Adjust the pixelSize as needed
        },
        "fields": "pixelSize"
    }
}
# Resize the a column to a specific width (e.g., 10 pixels)
resize_column_request_a_column = {
    "updateDimensionProperties": {
        "range": {
            "sheetId": num_id,  # Replace with the correct sheetId
            "dimension": "COLUMNS",
            "startIndex": 0,  # 0-based index of the F column
            "endIndex": 1
        },
        "properties": {
            "pixelSize": 150  # Adjust the pixelSize as needed
        },
        "fields": "pixelSize"
    }
}
# Resize the d column to a specific width (e.g., 10 pixels)
resize_column_request_d_column = {
    "updateDimensionProperties": {
        "range": {
            "sheetId": num_id,  # Replace with the correct sheetId
            "dimension": "COLUMNS",
            "startIndex": 3,  # 0-based index of the F column
            "endIndex": 4
        },
        "properties": {
            "pixelSize": 150  # Adjust the pixelSize as needed
        },
        "fields": "pixelSize"
    }
}
# Resize the B column to a specific width (e.g., 10 pixels)
resize_column_request_B_column = {
    "updateDimensionProperties": {
        "range": {
            "sheetId": num_id,  # Replace with the correct sheetId
            "dimension": "COLUMNS",
            "startIndex": 1,  # 0-based index of the F column
            "endIndex": 2
        },
        "properties": {
            "pixelSize": 135  # Adjust the pixelSize as needed
        },
        "fields": "pixelSize"
    }
}

# Resize the c column to a specific width (e.g., 10 pixels)
resize_column_request_c_column = {
    "updateDimensionProperties": {
        "range": {
            "sheetId": num_id,  # Replace with the correct sheetId
            "dimension": "COLUMNS",
            "startIndex": 2,  # 0-based index of the F column
            "endIndex": 3
        },
        "properties": {
            "pixelSize": 65  # Adjust the pixelSize as needed
        },
        "fields": "pixelSize"
    }
}








# Set the background color to red
background_color_request={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,
            "endRowIndex": 3,
            "startColumnIndex": 0,
            "endColumnIndex": 4
        },
        "cell": {
            "userEnteredFormat": {
                "backgroundColor": {"red": 0.7, "green": 0.7, "blue": 0.7}
            }
        },
        "fields": "userEnteredFormat(backgroundColor)"
    }
}

value_range_body={
    'majorDimension': 'ROWS',
    'values': values
}




# Apply data validation to column G
data_validation_request_for_g = {
    "setDataValidation": {
        "range": {
            "sheetId": num_id,  # Adjust the sheet ID as needed
            "startRowIndex": 4,  # Starting row index for column G
            "endRowIndex": 1000,  # Ending row index for column G (adjust as needed)
            "startColumnIndex": 6,  # Column G
            "endColumnIndex": 7
        },
        "rule": {
            "condition": {
                "type": "ONE_OF_LIST",
                "values": [{"userEnteredValue": str(i)} for i in range(21)]  # Numbers from 0 to 20
            },
            "showCustomUi": True,
            "inputMessage": "Select a number from 0 to 20"
        }
    }
}

# Apply data validation to column H
data_validation_request_for_h = {
    "setDataValidation": {
        "range": {
            "sheetId": num_id,  # Adjust the sheet ID as needed
            "startRowIndex": 4,  # Starting row index for column G
            "endRowIndex": 1000,  # Ending row index for column G (adjust as needed)
            "startColumnIndex": 7,  # Column G
            "endColumnIndex": 8
        },
        "rule": {
            "condition": {
                "type": "ONE_OF_LIST",
                "values": [{"userEnteredValue": str(i)} for i in range(21)]  # Numbers from 0 to 20
            },
            "showCustomUi": True,
            "inputMessage": "Select a number from 0 to 20"
        }
    }
}



#for A1 text
format_request={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,
            "endRowIndex": 1,
            "startColumnIndex": 0,
            "endColumnIndex": 1,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 17,
                    "bold":True,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            },
        },
        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)",
    }
}
# for A4 text
bold_text_request_A4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 0,  # Column A
            "endColumnIndex": 1,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "bold": True,
                    "fontSize": 10
                }
            }
        },
        "fields": "userEnteredFormat(textFormat)"
    }
}
# for B4 text
bold_text_request_B4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 1,  # Column A
            "endColumnIndex": 2,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "bold": True,
                    "fontSize": 10
                }
            }
        },
        "fields": "userEnteredFormat(textFormat)"
    }
}
# for C4 text
bold_text_request_C4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 2,  # Column A
            "endColumnIndex": 3,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "bold": True,
                    "fontSize": 10
                }
            }
        },
        "fields": "userEnteredFormat(textFormat)"
    }
}
# for D4 text
bold_text_request_D4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 3,  # Column A
            "endColumnIndex": 4,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "bold": True,
                    "fontSize": 10
                }
            }
        },
        "fields": "userEnteredFormat(textFormat)"
    }
}
# for e1-f1 text
bold_text_request_e1={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Row index for A4
            "endRowIndex": 1,
            "startColumnIndex": 4,  # Column A
            "endColumnIndex": 6,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold":True,
                },
                "backgroundColor": {
                    "red": 0.77,
                    "green": 0.77,
                    "blue": 0.77,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment)",
    }
}
#for e4 cell
bold_text_request_e4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 4,  # Column A
            "endColumnIndex": 5,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
#                "backgroundColor": {
#                    "red": 1.0,
#                    "green": 1.0,
#                    "blue": 0.0,
#                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)",
    }
}

#for f4 cell
bold_text_request_f4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 5,  # Column A
            "endColumnIndex": 6,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
#                "backgroundColor": {
#                    "red": 1.0,
#                    "green": 1.0,
#                    "blue": 0.0,
#                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)",
    }
}

#for G1 cell
bold_text_request_g1={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Row index for A4
            "endRowIndex": 1,
            "startColumnIndex": 6,  # Column A
            "endColumnIndex": 7,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.77,
                    "green": 0.77,
                    "blue": 0.77,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}


#for H1 cell
bold_text_request_h1={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Row index for A4
            "endRowIndex": 1,
            "startColumnIndex": 7,  # Column A
            "endColumnIndex": 8,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.77,
                    "green": 0.77,
                    "blue": 0.77,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}


#for e2 cell
bold_text_request_e2={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 1,  # Row index for A4
            "endRowIndex": 2,
            "startColumnIndex": 4,  # Column A
            "endColumnIndex": 5,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 14,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.7,
                    "green": 0.7,
                    "blue": 0.7,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}


#for G4 cell
bold_text_request_g4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 6,  # Column A
            "endColumnIndex": 7,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
#                "backgroundColor": {
#                    "red": 1.0,
#                    "green": 1.0,
#                    "blue": 0.0,
#                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)",
    }
}


#for H4 cell
bold_text_request_h4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 7,  # Column A
            "endColumnIndex": 8,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
#                "backgroundColor": {
#                    "red": 1.0,
#                    "green": 1.0,
#                    "blue": 0.0,
#                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)",
    }
}


#for G2 cell
bold_text_request_g2={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 1,  # Row index for A4
            "endRowIndex": 2,
            "startColumnIndex": 6,  # Column A
            "endColumnIndex": 7,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 14,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.7,
                    "green": 0.7,
                    "blue": 0.7,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}

#for h2 cell
bold_text_request_h2={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 1,  # Row index for A4
            "endRowIndex": 2,
            "startColumnIndex": 7,  # Column A
            "endColumnIndex": 8,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 14,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.7,
                    "green": 0.7,
                    "blue": 0.7,
               },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}


#for j1 cell


border_request_j1= {
    "updateBorders": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Row index for A4
            "endRowIndex": 1,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}


bold_text_request_j1={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Row index for A4
            "endRowIndex": 1,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.7,
                    "green": 0.7,
                    "blue": 0.7,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}

#for j2-j4 cell


border_request_j2_j4 = {
    "updateBorders": {
        "range": merge_range_j2_j4,
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}

bold_text_request_j2_j4={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 1,  # Row index for A4
            "endRowIndex": 4,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 18,
                },
                "backgroundColor": {
                    "red": 0.95,
                    "green": 0.95,
                    "blue": 0.95,
                },

                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment)",
    }
}


#for j5 cell
bold_text_request_j5={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 4,  # Row index for A4
            "endRowIndex": 5,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.95,
                    "green": 0.95,
                    "blue": 0.95,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}



#for j6 cell

border_request_j6= {
    "updateBorders": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 5,  # Row index for A4
            "endRowIndex": 6,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}

bold_text_request_j6={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 5,  # Row index for A4
            "endRowIndex": 6,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.7,
                    "green": 0.7,
                    "blue": 0.7,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}


#for j7j8 cell

border_request_j7j8 = {
    "updateBorders": {
        "range": merge_range_j7_j8,
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}

bold_text_request_j7_j8={
    "repeatCell": {
        "range": merge_range_j7_j8,
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 18,
                },
                "backgroundColor": {
                    "red": 0.95,
                    "green": 0.95,
                    "blue": 0.95,
                },

                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment)",
    }
}


#for j10 cell

bold_text_request_j10={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 9,  # Row index for A4
            "endRowIndex": 10,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.95,
                    "green": 0.95,
                    "blue": 0.95,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}



#for j11 cell

border_request_j11= {
    "updateBorders": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 10,  # Row index for A4
            "endRowIndex": 11,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}

bold_text_request_j11={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 10,  # Row index for A4
            "endRowIndex": 11,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.7,
                    "green": 0.7,
                    "blue": 0.7,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}









#for j12-j14 cell

border_request_j12_j14= {
    "updateBorders": {
        "range": merge_range_j12_j14,
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}

bold_text_request_j12_j14={
    "repeatCell": {
        "range":merge_range_j12_j14,
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 18,

                },
                "backgroundColor": {
                    "red": 0.95,
                    "green": 0.95,
                    "blue": 0.95,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}



#for j15 cell

bold_text_request_j15={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 14,  # Row index for A4
            "endRowIndex": 15,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.95,
                    "green": 0.95,
                    "blue": 0.95,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}


#for j11 cell

border_request_j11= {
    "updateBorders": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 10,  # Row index for A4
            "endRowIndex": 11,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}


#for j16 cell

border_request_j16= {
    "updateBorders": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 15,  # Row index for A4
            "endRowIndex": 16,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}

bold_text_request_j16={
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 15,  # Row index for A4
            "endRowIndex": 16,
            "startColumnIndex": 9,  # Column A
            "endColumnIndex": 10,
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 10,
                    "bold": True,
                },
                "backgroundColor": {
                    "red": 0.7,
                    "green": 0.7,
                    "blue": 0.7,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}


#for j17-j19 cell

border_request_j17_j19= {
    "updateBorders": {
        "range": merge_range_j17_j19,
        "top": border_style,
        "bottom": border_style,
        "left": border_style,
        "right": border_style,
    }
}

bold_text_request_j17_j19={
    "repeatCell": {
        "range":merge_range_j17_j19,
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontSize": 18,

                },
                "backgroundColor": {
                    "red": 0.95,
                    "green": 0.95,
                    "blue": 0.95,
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }
        },

        "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)",
    }
}






# Define the range for the entire column I
format_range_column_i = {
    "sheetId": num_id,
    "startRowIndex": 0,  # Starting from row 4
    "endRowIndex": 1000,  # Ending at row 1000 (adjust as needed)
    "startColumnIndex": 8,  # Column I
    "endColumnIndex": 9  # End column index (just I)
}

# Create a request to set the background color for the entire column I
background_color_request_column_i = {
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [format_range_column_i],
            "booleanRule": {
                "condition": {
                    "type": "CUSTOM_FORMULA",
                    "values": [{"userEnteredValue": "=TRUE"}]
                },
                "format": {
                    "backgroundColor": {
                        "red": 0.95,  # Adjust these values for the desired color (e.g., grey)
                        "green": 0.95,
                        "blue": 0.95,
                    }
                },
            },
        },
        "index": 0,
    }
}


# Define the range for the entire column K
format_range_column_k = {
    "sheetId": num_id,
    "startRowIndex": 0,  # Starting from row 4
    "endRowIndex": 1000,  # Ending at row 1000 (adjust as needed)
    "startColumnIndex": 10,  # Column K
    "endColumnIndex": 11  # End column index (just K)
}

# Create a request to set the background color for the entire column K
background_color_request_column_k = {
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [format_range_column_k],
            "booleanRule": {
                "condition": {
                    "type": "CUSTOM_FORMULA",
                    "values": [{"userEnteredValue": "=TRUE"}]
                },
                "format": {
                    "backgroundColor": {
                        "red": 0.95,  # Adjust these values for the desired color (e.g., grey)
                        "green": 0.95,
                        "blue": 0.95,
                    }
                },
            },
        },
        "index": 0,
    }
}


background_color_row4 = {
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 3,  # Row index for row 4 (0-based index)
            "endRowIndex": 4,    # Row index for row 4 (0-based index)
            "startColumnIndex": 0,  # Start from the first column
            "endColumnIndex": 26,  # Adjust to cover all columns you need (A-Z)
        },
        "cell": {
            "userEnteredFormat": {
                "backgroundColor": {
                    "red": 0.77,  # Adjust these values for yellow color
                    "green": 0.77,
                    "blue": 0.77,
                }
            }
        },
        "fields": "userEnteredFormat.backgroundColor"
    }
}

# Data validation for column E (with checked and unchecked checkboxes)
data_validation_request_e = {
    "setDataValidation": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 4,  # Starting row index for column E
            "endRowIndex": 1000,  # Ending row index for column E (adjust as needed)
            "startColumnIndex": 4,  # Column E
            "endColumnIndex": 5
        },
        "rule": {
            "condition": {
                "type": "BOOLEAN",
            },
            "showCustomUi": True,
            "inputMessage": "Click the cell to toggle checkbox",
        }
    }
}
# Conditional formatting request to change cell background color to green when checkbox is marked
conditional_formatting_request = {
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [
                {
                    "sheetId": num_id,  # Sheet ID
                    "startRowIndex": 3,  # Start row index
                    "endRowIndex": 1000,  # End row index
                    "startColumnIndex": 4,  # Column E
                    "endColumnIndex": 6,
                }
            ],
            "booleanRule": {
                "condition": {
                    "type": "CUSTOM_FORMULA",
                    "values": [{"userEnteredValue": '=E4=TRUE'}]  # Adjust row number as needed
                },
                "format": {
                    "backgroundColor": {
                        "red": 0.0,
                        "green": 1.0,  # Set the green color to maximum
                        "blue": 0.0,
                    }
                },
            }
        },
        "index": 0,
    }
}



#for column B namberFormat
req_for_colomn_b = {
    'repeatCell': {
        'range': {
            'sheetId': num_id,
            'startRowIndex': 1,  # Adjust as needed
            'endRowIndex': 100,  # Adjust as needed
            'startColumnIndex': 1,  # Column B
            'endColumnIndex': 2,
        },
        'cell': {
            'userEnteredFormat': {
                'numberFormat': {
                    'type': 'TEXT',
                },
            },
        },
        'fields': 'userEnteredFormat.numberFormat',
    },
}



column_letter='B5:B'
# Define the data validation request for numeric input in column B using a custom formula
data_validation_request_b = {
    'setDataValidation': {
        'range': {
            'sheetId': num_id,  # Replace with the sheet ID
            'startRowIndex': 4,  # Starting row index for column B
            'endRowIndex': 1000,  # Ending row index for column B (adjust as needed)
            'startColumnIndex': 1,  # Column B
            'endColumnIndex': 2
        },
        'rule': {
            "condition": {
        "type": "CUSTOM_FORMULA",
        "values": [{
            "userEnteredValue": f'=AND(LEN({column_letter}5:{column_letter}) = 10, LEFT({column_letter}5:{column_letter}, 2) = "05", ISNUMBER(VALUE({column_letter}5:{column_letter})), OR(MID({column_letter}5:{column_letter}, 3, 1) = "0", MID({column_letter}5:{column_letter}, 3, 1) = "2", MID({column_letter}5:{column_letter}, 3, 1) = "3", MID({column_letter}5:{column_letter}, 3, 1) = "4", MID({column_letter}5:{column_letter}, 3, 1) = "8"))',
        }],

    },
            'inputMessage': 'מספר טלפון לא תקין',
            'strict':False,

            'showCustomUi': True,  # Set to True to show a custom error message


        }
    }


}








# Data validation for column f (with checked and unchecked checkboxes)
data_validation_request_f = {
    "setDataValidation": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 4,  # Starting row index for column E
            "endRowIndex": 1000,  # Ending row index for column E (adjust as needed)
            "startColumnIndex": 5,  # Column E
            "endColumnIndex": 6
        },
        "rule": {
            "condition": {
                "type": "BOOLEAN",
            },
            "showCustomUi": True,
            "inputMessage": "Click the cell to toggle checkbox",
        }
    }
}


# Calculate the sum in G2
calculate_sum_request_g2 = {
    "updateCells": {
        "rows": [
            {
                "values": [
                    {
                        "userEnteredValue": {
                            "formulaValue": "=SUM(G4:G1000)"
                        }
                    }
                ]
            }
        ],
        "fields": "userEnteredValue.formulaValue",
        "start": {
            "sheetId": num_id,  # Adjust the sheet ID as needed
            "rowIndex": 1,  # Row index for G2
            "columnIndex": 6  # Column G
        }
    }
}


# Calculate the sum in h2
calculate_sum_request_h2 = {
    "updateCells": {
        "rows": [
            {
                "values": [
                    {
                        "userEnteredValue": {
                            "formulaValue": "=SUM(H4:H1000)"
                        }
                    }
                ]
            }
        ],
        "fields": "userEnteredValue.formulaValue",
        "start": {
            "sheetId": num_id,  # Adjust the sheet ID as needed
            "rowIndex": 1,  # Row index for G2
            "columnIndex": 7  # Column G
        }
    }
}


# Define alternating row colors
alternating_row = {
    "addConditionalFormatRule": {
        "rule": {
            "ranges": [
                {
                    "sheetId": num_id,  # Adjust the sheet ID as needed
                    "startRowIndex": 4,  # Start from row 4
                    "endRowIndex": 1000,
                    "startColumnIndex":0,
                    "endColumnIndex":8,
                },
            ],
            "booleanRule": {
                "condition": {
                    "type": "CUSTOM_FORMULA",
                    "values": [{"userEnteredValue": '=ISEVEN(ROW())'}]
                },
                "format": {
                    "backgroundColor": {
                        "red": 0.9,  # Adjust these values for the desired color
                        "green": 0.9,
                        "blue": 0.9,
                    }
                },
            },
        },
        "index": 0,
    }
}


text_format_request_to_varela = {
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Start from the first row
            "endRowIndex": 1000,  # Adjust to cover all rows you need
            "startColumnIndex": 0,  # Start from the first column
            "endColumnIndex": 26,  # Adjust to cover all columns you need (A-Z)
        },
        "cell": {
            "userEnteredFormat": {
                "textFormat": {
                    "fontFamily": "Varela Round",  # Set the desired font family
                }
            }
        },
        "fields": "userEnteredFormat.textFormat.fontFamily"
    }
}


#alignment_request_for_a
alignment_request_for_a = {
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Adjust the row index as needed
            "endRowIndex": 1048576,  # Maximum number of rows
            "startColumnIndex": 0,  # Adjust the column index as needed
            "endColumnIndex": 1,  # Maximum number of columns
        },
        "cell": {
            "userEnteredFormat": {
                "horizontalAlignment": "RIGHT"  # Set horizontal alignment to right
            }
        },
        "fields": "userEnteredFormat.horizontalAlignment"
    }
}

#alignment_request_for_c
alignment_request_for_c = {
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Adjust the row index as needed
            "endRowIndex": 1048576,  # Maximum number of rows
            "startColumnIndex": 2,  # Adjust the column index as needed
            "endColumnIndex": 3,  # Maximum number of columns
        },
        "cell": {
            "userEnteredFormat": {
                "horizontalAlignment": "RIGHT"  # Set horizontal alignment to right
            }
        },
        "fields": "userEnteredFormat.horizontalAlignment"
    }
}
#alignment_request_for_d
alignment_request_for_d = {
    "repeatCell": {
        "range": {
            "sheetId": num_id,
            "startRowIndex": 0,  # Adjust the row index as needed
            "endRowIndex": 1048576,  # Maximum number of rows
            "startColumnIndex": 3,  # Adjust the column index as needed
            "endColumnIndex": 4,  # Maximum number of columns
        },
        "cell": {
            "userEnteredFormat": {
                "horizontalAlignment": "RIGHT"  # Set horizontal alignment to right
            }
        },
        "fields": "userEnteredFormat.horizontalAlignment"
    }
}





# Define the start and end indexes for columns L to Z
start_column_index = 11  # 0-based index for column L
end_column_index = 26   # 0-based index for column Z

# Create a request to delete columns L to Z
delete_columns_request = {
    "deleteDimension": {
        "range": {
            "sheetId": num_id,
            "dimension": "COLUMNS",
            "startIndex": start_column_index,
            "endIndex": end_column_index + 1  # Add 1 to include the end column
        }
    }
}

# Batch update requests
batch_update_request={
    "requests": [merge_request_e1_f1, merge_request, merge_request_e2_f3, alternating_row,
                 merge_request_g2_g3, merge_request_h2_h3, merge_request_i2_i10000, text_direction_request,
                 data_validation_request_f, bold_text_request_g2, conditional_formatting_request, req_for_colomn_b,
                 background_color_row4,alignment_request_for_a,alignment_request_for_c,alignment_request_for_d,
                 resize_column_request_a_column, resize_column_request_d_column, resize_column_request_B_column,
                 resize_column_request_c_column,data_validation_request_b,
                 resize_column_request_i_column, bold_text_request_h4, data_validation_request_for_g,
                 merge_request_j7_j8, merge_request_j12_j14, merge_request_j17_j19,
                 data_validation_request_for_h, calculate_sum_request_h2, resize_column_request_j_column,
                 resize_column_request_k_column, delete_columns_request, merge_request_k2_k10000,
                 merge_request_j2_j4, border_request_j2_j4, border_request_j1, border_request_j6,
                 format_request, bold_text_request_A4, bold_text_request_B4, bold_text_request_f4,
                 bold_text_request_C4, bold_text_request_D4, bold_text_request_e2, bold_text_request_j5,
                 bold_text_request_e1, bold_text_request_e4, bold_text_request_g4, bold_text_request_j2_j4,
                 bold_text_request_j7_j8, bold_text_request_j10, bold_text_request_j11, bold_text_request_j12_j14,
                 bold_text_request_j15, bold_text_request_j16, bold_text_request_j17_j19, border_request_j17_j19,
                 border_request_j7j8, border_request_j11, border_request_j12_j14, border_request_j16,
                 bold_text_request_h2, bold_text_request_j1, background_color_request,
                 background_color_request_column_i, bold_text_request_j6, background_color_request_column_k
                 , bold_text_request_g1, bold_text_request_h1, data_validation_request_e, calculate_sum_request_g2,
                 text_format_request_to_varela]
}

# Execute the batch update request
request=service.spreadsheets().batchUpdate(spreadsheetId=Id, body=batch_update_request)
response=request.execute()




values_a4=[
    ['שם מלא'],
]
cell_range_a4='A4'

values_b4=[
    ["מס' פלאפון"],
]
cell_range_b4='B4'

values_c4=[
    ["צד"],
]
cell_range_c4='C4'

values_d4=[
    ["קבוצה"],
]
cell_range_d4='D4'

values_e1=[
    ["% אישרו הגעה "],
]
cell_range_e1='E1'

values_e4=[
    ['נשלחה הודעה'],
]
cell_range_e4='E4'

values_f4=[
    ['אישרו הגעה'],
]
cell_range_f4='f4'

values_g1=[
    ['סך מוזמנים'],
]
cell_range_g1='G1'

values_h1=[
    ['סך אישורים'],
]
cell_range_h1='H1'

values_g4=[
    ['מוזמנים'],
]
cell_range_g4='g4'

values_h4=[
    ["מגיעים"],
]
cell_range_h4='H4'

values_j1=[
    ['תאריך'],
]
cell_range_j1='J1'

values_j6=[
    ['ימים שנשארו'],
]
cell_range_j6='J6'

values_j11=[
    ["הזמנות שנשלחו"],
]
cell_range_j11='J11'

values_j16=[
    ["הזמנות שנשארו לשלוח"],
]
cell_range_j16='J16'

values_j17=[
    ['=IF(COUNTA(A5:A) - COUNTIF(E5:E, TRUE) > 0, COUNTA(A5:A) - COUNTIF(E5:E, TRUE), "Done!")'],
]
cell_range_j17='J17'


formula_e = '=COUNTIF(E3:E, TRUE)'
formula_f = '=COUNTIF(F4:F, TRUE)'

formula_e2 = '=IFERROR(TEXT(MIN((COUNTIF(F3:F, TRUE) / COUNTIF(E4:E, TRUE)) * 100, 100), "0.00") & "%", "")'

values_j12=[
    [formula_e],
]
cell_range_j12='J12'


# Values and cell range for cell E2
values_e2 = [[formula_e2]]
cell_range_e2 = 'E2'

















service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_insert,
    body=value_range_body

).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_a4,
    body={'majorDimension': 'ROWS', 'values': values_a4}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_b4,
    body={'majorDimension': 'ROWS', 'values': values_b4}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_c4,
    body={'majorDimension': 'ROWS', 'values': values_c4}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_d4,
    body={'majorDimension': 'ROWS', 'values': values_d4}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_e1,
    body={'majorDimension': 'ROWS', 'values': values_e1}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_e4,
    body={'majorDimension': 'ROWS', 'values': values_e4}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_f4,
    body={'majorDimension': 'ROWS', 'values': values_f4}
).execute()



service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_g1,
    body={'majorDimension': 'ROWS', 'values': values_g1}
).execute()


service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_h1,
    body={'majorDimension': 'ROWS', 'values': values_h1}
).execute()


service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_e2,
    body={'majorDimension': 'ROWS', 'values': values_e2}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_g4,
    body={'majorDimension': 'ROWS', 'values': values_g4}
).execute()

service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_h4,
    body={'majorDimension': 'ROWS', 'values': values_h4}
).execute()



service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_j1,
    body={'majorDimension': 'ROWS', 'values': values_j1}
).execute()



service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_j6,
    body={'majorDimension': 'ROWS', 'values': values_j6}
).execute()


service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_j12,
    body={'majorDimension': 'ROWS', 'values': values_j12}
).execute()


service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_j11,
    body={'majorDimension': 'ROWS', 'values': values_j11}
).execute()



service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_j16,
    body={'majorDimension': 'ROWS', 'values': values_j16}
).execute()



service.spreadsheets().values().update(
    spreadsheetId=Id,
    valueInputOption='USER_ENTERED',
    range=work_sheet_name + cell_range_j17,
    body={'majorDimension': 'ROWS', 'values': values_j17}
).execute()




end_time = time.time()
elapsed_time = end_time - start_time
print(f"Elapsed time: {elapsed_time:.4f} seconds")