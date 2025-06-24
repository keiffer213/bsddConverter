import os
import json
import numpy as np
import pandas as pd
from ast import literal_eval
from copy import deepcopy
from tqdm import tqdm
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def run_excel2bsdd_conversion(excel_path, template_path, output_path, remove_nulls=False):
    """
    Main function to map excel file to bsdd json file

    :param excel_path: Path to an excel file 
    :type excel_path: str
    :param template_path: Path to JSON template file
    :type template_path: str
    :param output_path: Path to output JSON file
    :type output_path: str
    """

    # Check if files exist
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template file not found: {template_path}")

    # Load files
    with open(template_path, encoding="utf-8") as f:
        tpl = json.load(f)

    excel = load_excel(excel_path)
    result = excel2bsdd(excel, tpl)

    # Save file
    if remove_nulls:
        result = clean_nones(result)
    if result is None:
        raise ValueError("Conversion result is None. Please check input files.")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

def load_excel(EXCEL_PATH):
    """
    Parses an excel file from path. Note: only works on provided template file.

    :param EXCEL_PATH: Path to an excel file 
    :type EXCEL_PATH: str
    :return: Dictionary of Pandas dataframes with parsed Excel data
    :rtype: dict
    """

    try:
        excel_df = pd.ExcelFile(EXCEL_PATH)
    except PermissionError:
        raise Exception("Excel file is open. Please close it and try again.")
    
    excel = {}
    excel['dictionary'] = pd.read_excel(excel_df, 'Dictionary', skiprows=6, usecols="C:R", true_values="TRUE", keep_default_na=False, converters={'DictionaryVersion': str})    
    excel['class'] = pd.read_excel(excel_df, 'Class', skiprows=6, usecols="C:AC", true_values="TRUE", keep_default_na=False, converters={'Uid': str})
    excel['property'] = pd.read_excel(excel_df, 'Property', skiprows=6, usecols="C:AU", true_values="TRUE", keep_default_na=False)   
    excel['classproperty'] = pd.read_excel(excel_df, 'ClassProperty', usecols="C:U", skiprows=6, true_values="TRUE", keep_default_na=False)
    excel['classrelation'] = pd.read_excel(excel_df, 'ClassRelation', usecols="C:H", skiprows=6, true_values="TRUE", keep_default_na=False)
    excel['allowedvalue'] = pd.read_excel(excel_df, 'AllowedValue', skiprows=6, usecols="C:J", true_values="TRUE", keep_default_na=False) 
    excel['propertyrelation'] = pd.read_excel(excel_df, 'PropertyRelation', skiprows=6, usecols="C:G", true_values="TRUE", keep_default_na=False)
    return excel

def map_data(excel_data, bsdd_part_template, name=""):
    """
    Transforms the input pandas dataframe to JSON only if a property exists in the template

    :param excel_data: Pandas dataframe with parsed Excel data
    :type excel_data: pd.DataFrame
    :param json_part: template dictinary from JSON_templates
    :type json_part: dict
    :return: Resultant list of dictionaries containing each row of the pandas table converted to appropriate dictionary
    :rtype: list
    """

    if isinstance(bsdd_part_template, list):
        template = deepcopy(bsdd_part_template[0])
    else:
        template = deepcopy(bsdd_part_template)

    for k, v in template.items():
        if isinstance(v, list):
            template[k] = []

    excel_data = excel_data.replace(r'^\s*$', np.nan, regex=True)
    excel_data = excel_data.astype(object).replace(np.nan, None)
    new_objects = []

    for index, row in tqdm(excel_data.iterrows(), desc=f"Processing {name}", unit=" items", total=len(excel_data)):
        if not excel_data.dropna(how="all").empty:
            new_object = deepcopy(template)
            for column_name, column_data in row.items():
                if column_name in template:
                    # Convert date to: 2022-05-12T00:00:00+02:00
                    if isinstance(column_data, pd._libs.tslibs.timestamps.Timestamp):
                        column_data = column_data.isoformat()
                    elif "Date" in column_name and column_data:
                        column_data = pd.to_datetime(column_data, origin='1899-12-30', unit='D').isoformat()
                    elif (column_name in ["RevisionNumber","VersionNumber","SortNumber"] or (column_name[0:9]=="Dimension" and len(column_name)>9)) and column_data is not None:
                        column_data = int(column_data)
                    elif column_name in ["Uid","Example","Value","PredefinedValue"] and not isinstance(column_data, str):
                        column_data = str(column_data)

                    # Process lists
                    if isinstance(column_data, str):
                        if column_data.startswith("[") and column_data.endswith("]"):
                            if column_data == "[]":
                                column_data = None
                            else:
                                content = literal_eval(column_data)
                                if isinstance(content, list):
                                    column_data = content
                    if column_name in ["RelatedIfcEntityNamesList","Units","ReplacedObjectCodes","ReplacingObjectCodes","CountriesOfUse","SubdivisionsOfUse"]:
                        if not isinstance(column_data, list):
                            column_data = [column_data]
                    # append
                    new_object[column_name] = column_data
                elif column_name in ('(Origin Class Code)','(Origin Property Code)','(Origin ClassProperty Code)'):
                    new_object[column_name] = column_data
                else:
                    print(f"WARNING! No such property as '{column_name}' in the JSON template! It will NOT be added to the JSON file.")
                    # new_object[column_name] = column_data
            # bsdd_part_template.append(new_object)
            new_objects.append(new_object)
    return new_objects

def clean_nones(value):
    """
    Recursively remove all None values from dictionaries and lists, and returns
    the result as a new dictionary or list.
    """

    if isinstance(value, list):
        return [clean_nones(x) for x in value if x not in ("", [], None)]
    elif isinstance(value, dict):
        return {
            key: clean_nones(val)
            for key, val in value.items()
            if val not in ("", [], None)
        }
    else:
        return value

def excel2bsdd(excel, bsdd_template):
    """
    Goes through all dataframes and appends data to the desired JSON structure

    :param excel: Dictionary of Pandas dataframes from load_excel
    :type excel: dict
    :return: Resultant JSON structure
    :rtype: dict
    """

    bsdd_data = map_data(excel['dictionary'], bsdd_template, "dictionary")[0]

    # process basic concepts
    bsdd_data['Classes'] = map_data(excel['class'], bsdd_template['Classes'], "classes")
    bsdd_data['Properties'] = map_data(excel['property'], bsdd_template['Properties'], "properties")

    # process ClassProperty
    cls_props = map_data(excel['classproperty'], bsdd_template['Classes'][0]['ClassProperties'], "class-properties")
    for cls_prop in cls_props:
        related = cls_prop['(Origin Class Code)']
        cls_prop.pop("(Origin Class Code)")
        found_it = False
        for item in bsdd_data['Classes']:
            if item["Code"] == related:
                item['ClassProperties'].append(cls_prop)
                found_it = True
                break
        if not found_it:
            raise Exception(f"Class '{related}' not found in the spreadsheet, so couldn't append the class property: '{cls_prop}'!")

    # process ClassRelation
    cls_rels = map_data(excel['classrelation'], bsdd_template['Classes'][0]['ClassRelations'], "class-relations")
    for cls_rel in cls_rels:
        related = cls_rel['(Origin Class Code)']
        cls_rel.pop("(Origin Class Code)")
        found_it = False
        for item in bsdd_data['Classes']:
            if item["Code"] == related:
                item['ClassRelations'].append(cls_rel)
                found_it = True
                break
        if not found_it:
            raise Exception(f"Class '{related}' not found in the spreadsheet, so couldn't append the class relation: '{cls_rel}'!")

    # process AllowedValue
    allowed_vals = map_data(excel['allowedvalue'], bsdd_template['Properties'][0]['AllowedValues'], "allowed-values")
    for allowed_val in allowed_vals:
        # Only one of two code columns is possible
        if allowed_val['(Origin Property Code)']:
            relToProperty = True
            related = allowed_val['(Origin Property Code)']
        elif allowed_val['(Origin ClassProperty Code)']:
            relToProperty = False
            related = allowed_val['(Origin ClassProperty Code)']
        else:
            print("WARNING! Allowed value without origin property or classProperty code! It will NOT be added to the JSON file.")
        allowed_val.pop("(Origin Property Code)")
        allowed_val.pop("(Origin ClassProperty Code)")
        if relToProperty:
            # iterate all properties and add AllowedValue if such property is present in the spreadsheet
            found_it = False
            for item in bsdd_data['Properties']:
                if item['Code'] == related:
                    item['AllowedValues'].append(allowed_val)
                    found_it = True
                    break
            if not found_it:
                raise Exception(f"Property '{related}' not found in the spreadsheet, so couldn't append the e {allowed_val}!")
        else:
            # iterate all classes to find the one referenced by the property AllowedValue
            found_it = False
            for cl in bsdd_data['Classes']:
                for item in cl['ClassProperties']:
                    if item["Code"] == related:
                        item['AllowedValues'].append(allowed_val)
                        found_it = True
                        break
            if not found_it:
                raise Exception(f"Class '{related}' not found in the spreadsheet, so couldn't append the e {allowed_val}!")

    # process PropertyRelation
    prop_rels = map_data(excel['propertyrelation'], bsdd_template['Properties'][0]['PropertyRelations'], "property-relations")
    for prop_rel in prop_rels:
        related = prop_rel['(Origin Property Code)']
        prop_rel.pop("(Origin Property Code)")
        found_it = False
        for item in bsdd_data['Properties']:
            if item["Code"] == related:
                item['PropertyRelations'].append(prop_rel)
                found_it = True
                break
        if not found_it:
            raise Exception(f"Class '{related}' not found in the spreadsheet, so couldn't append the value {prop_rel}!")

    return bsdd_data