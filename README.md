# task
import pandas as pd
import json
import logging


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def parse_worksheet(sheet_name, sheet_df):
    course_data= {
        "course_code": None,
        "course_title": None,
        "credit_structure": {},
        "sections":[]
    }
    
    
        
    course_data["course_code"] = str(sheet_df.iloc[2, 1]) if len(sheet_df.columns) > 1 else None
    course_data["course_title"] = str(sheet_df.iloc[2, 2]) if len(sheet_df.columns) > 2 else None
        
        
    course_data["credit_structure"] = {
            "L": int(sheet_df.iloc[2, 3]) if pd.notna(sheet_df.iloc[2, 3]) and isinstance(sheet_df.iloc[2, 3], (int, float)) else 0,
            "P": int(sheet_df.iloc[2, 4]) if pd.notna(sheet_df.iloc[2, 4]) and isinstance(sheet_df.iloc[2, 4], (int, float)) else 0,
            "U": int(sheet_df.iloc[2, 5]) if pd.notna(sheet_df.iloc[2, 5]) and isinstance(sheet_df.iloc[2, 5], (int, float)) else 0
        }
        
    
    for i in range(2,len(sheet_df)):
        dict_i={
                    "section_type":(str(sheet_df.iloc[i][6]))[0],
                    "section number":(sheet_df.iloc[i][6]),
                    "instructors":str(sheet_df.iloc[i][7]),
                    "room":sheet_df.iloc[i][8],
                    "time_slots":str(sheet_df.iloc[i][9])

        }
        course_data["sections"].append(dict_i)
        
        print(course_data["sections"])
    return(course_data)       
            

def main():
    
    file_path = '/Users/aditya/Downloads/Timetable Workbook - SUTT Task 1.xlsx' 
    workbook = pd.ExcelFile(file_path)
    timetable_data = []

    # Process each sheet
    for sheet_name in workbook.sheet_names:
        sheet_df = workbook.parse(sheet_name)
        logging.info(f"Processing sheet: {sheet_name}")
        course_data = parse_worksheet(sheet_name, sheet_df)
        timetable_data.append(course_data)

    # Save the extracted data to a JSON file
    output_file = "/Users/aditya/Downloads/fileto.json"
    with open(output_file, "w") as json_file:
        json.dump(timetable_data, json_file, indent=4)

    logging.info(f"JSON data successfully saved to {output_file}")

if __name__ == "__main__":
    main()
