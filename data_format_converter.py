import pandas as pd
import json
import xml.etree.ElementTree as ET
import sqlite3
import os

# available output formats
available_formats = ['csv', 'json', 'xlsx', 'xml', 'sql', 'db']

def create_input_folder():
    folder_name = 'files'
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    return folder_name

def prompt_to_place_files(input_folder):
    print(f"Please place your input files into the '{input_folder}' folder.")
    input("Press Enter when ready to continue...")

def get_input_file_path(input_folder):
    while True:
        file_list = os.listdir(input_folder)
        if not file_list:
            print(f"No files found in '{input_folder}' folder. Please place your input files there and try again.")
            prompt_to_place_files(input_folder)
        else:
            print("Found the following files in the 'files' folder:")
            for i, file in enumerate(file_list):
                print(f"{i + 1}. {file}")
            
            file_index = input("Enter the number of the file you want to convert (or 'quit' to exit): ").strip()
            
            if file_index.lower() == 'quit':
                print("Exiting the converter.")
                return None
            elif file_index.isdigit() and 1 <= int(file_index) <= len(file_list):
                selected_file = file_list[int(file_index) - 1]
                return os.path.join(input_folder, selected_file)
            else:
                print("Invalid input. Please enter a valid file number or 'quit' to exit.")

def read_csv(file_path, encodings=('utf-8', 'latin-1', 'iso-8859-1', 'cp1252')):
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            return df
        except UnicodeDecodeError:
            print(f"Error reading CSV file '{file_path}' with encoding '{encoding}': UnicodeDecodeError. Trying next encoding...")
        except Exception as e:
            print(f"Error reading CSV file '{file_path}' with encoding '{encoding}': {str(e)}")
    print(f"Failed to read CSV file '{file_path}' with all specified encodings: {encodings}")
    return None

def read_json(file_path):
    with open(file_path, 'r') as file:
        data = json.load(file)
    return pd.json_normalize(data)

def read_excel(file_path):
    return pd.read_excel(file_path)

def read_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    def xml_to_dict(element):
        return {
            child.tag: xml_to_dict(child) if len(child) else child.text
            for child in element
        }

    data_dict = xml_to_dict(root)
    return pd.json_normalize(data_dict)

def read_sql(file_path, table_name='data'):
    conn = sqlite3.connect(file_path)
    query = f"SELECT * FROM {table_name};"
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def read_file(file_path, encoding='utf-8'):
    file_extension = file_path.split('.')[-1].lower()
    read_functions = {
        'csv': lambda path: read_csv(path, encodings=(encoding, 'latin-1', 'iso-8859-1', 'cp1252')),
        'json': read_json,
        'xlsx': read_excel,
        'xls': read_excel,
        'xml': read_xml,
        'db': read_sql,
        'sql': read_sql
    }

    if file_extension in read_functions:
        return read_functions[file_extension](file_path)
    else:
        raise ValueError('Unsupported file format')

def write_csv(dataframe, file_path):
    dataframe.to_csv(file_path, index=False)

def write_json(dataframe, file_path):
    data = dataframe.to_dict(orient='records')
    with open(file_path, 'w') as file:
        json.dump(data, file, indent=4)

def write_excel(dataframe, file_path):
    dataframe.to_excel(file_path, index=False)

def write_xml(dataframe, file_path):
    def dict_to_xml(tag, d):
        elem = ET.Element(tag)
        for key, val in d.items():
            child = ET.Element(key)
            child.text = str(val)
            elem.append(child)
        return elem
    
    root = dict_to_xml('root', dataframe.to_dict(orient='records')[0])
    tree = ET.ElementTree(root)
    tree.write(file_path)

def write_sql(dataframe, file_path, table_name='data'):
    conn = sqlite3.connect(file_path)
    dataframe.to_sql(table_name, conn, index=False, if_exists='replace')
    conn.commit()
    conn.close()

def write_file(dataframe, file_path, output_format):
    write_functions = {
        'csv': write_csv,
        'json': write_json,
        'xlsx': write_excel,
        'xml': write_xml,
        'db': lambda df, path: write_sql(df, path, table_name='data'),
        'sql': lambda df, path: write_sql(df, path, table_name='data')
    }

    if output_format in write_functions:
        write_functions[output_format](dataframe, file_path)
    else:
        raise ValueError('Unsupported output format')

def write_all_files(dataframe, base_file_path):
    write_functions = {
        'csv': write_csv,
        'json': write_json,
        'xlsx': write_excel,
        'xml': write_xml,
        'db': lambda df, path: write_sql(df, path, table_name='data'),
        'sql': lambda df, path: write_sql(df, path, table_name='data')
    }

    for format, write_function in write_functions.items():
        output_path = f"{base_file_path}.{format}"
        write_function(dataframe, output_path)
        print(f"Written to {output_path}")


if __name__ == "__main__":
    input_folder = create_input_folder()
    prompt_to_place_files(input_folder)
    
    while True:
        input_path = get_input_file_path(input_folder)
        if not input_path:
            break
        
        while True:
            print("\nAvailable formats for conversion:")
            for fmt in available_formats:
                print(f"- {fmt}")
            output_format = input("Enter the output file format (or 'all' to convert to all formats, or 'quit' to exit): ").lower().strip()
            
            if output_format == 'quit':
                print("Exiting the converter.")
                break
            elif output_format == 'all':
                base_output_path = input("Enter the base output file path (e.g., output): ").strip()
                
                try:
                    dataframe = read_file(input_path)
                except Exception as e:
                    print(f"Error reading input file: {str(e)}")
                    continue
                
                write_all_files(dataframe, base_output_path)
            elif output_format in available_formats:
                
                try:
                    dataframe = read_file(input_path)
                except Exception as e:
                    print(f"Error reading input file: {str(e)}")
                    continue
                
                output_path = f"{input_path.rsplit('.', 1)[0]}.{output_format}"
                write_file(dataframe, output_path, output_format)
                print(f"Written to {output_path}")
            else:
                print("Unsupported output format. Please choose a valid format.")
        
        continue_prompt = input("Do you want to convert another file? (yes/no): ").lower().strip()
        if continue_prompt != 'yes':
            print("Exiting the converter.")
            break
