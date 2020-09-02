
# Author    : Qiaojie Zheng, zheng@mymail.mines.edu, qiaojiez@andrew.cmu.edu
# Date      : 08/30/2020

# This code is used for matching image files with the collected data, and
# store the image file name alongside with the collected data. Data comes 
# from Colorado School of Mines and Carnegie Mellon University for Wire-
# Feed laser AM. 
import xlrd
from SheetData import *
from SchoolDataClass import *
from openpyxl.reader.excel import load_workbook



reportPath = "Citrination_export.xlsx"


def hasNumber(inputString):
    """ Checks if the input string contains a number
    This methods is used to check if the cell value is a valid sample same,
    sample names always contains numbers
    :param inputString: a string variable
    :returns : True if there is(are) number(s) in the input string
    """
    return any(char.isdigit() for char in inputString)

def create_comp_dict(metrology_dict, metallography_dict):
    
    """ Combines the Metrology and Metallography into a single dictionary.
    This method will discard entries in either one of the two dictionary that
    have empty data with the same parent sample name in the other one
    
    :param metrology_dict: a dictionary that contains metrology information, keys
                           the sample names, and the elements are Metrology class
                           variables defined in SchoolDataClass.py
                           
    :param metallography_dict: a dictionary that contains metallography information,
                               keys are the sample names, and elements are
                               Metallography class variables defined in SchoolDataClass.py
    :returns: a dictionary with keys being the parent sample names, and elements are 
              SampleData class found in the SchoolDataClass.py
    """
    
    toReturn = {}
    for key in metallography_dict.keys():
        
        if key in metrology_dict.keys(): # if the metrology dictionary has the same key in it
            curr_sample_data = SampleData(metrology_dict[key], metallography_dict[key]);
            toReturn[key] = curr_sample_data;

    return toReturn

def findMatchedResults(schoolData, sheetData):
    """ Searches for data entries that are in both shcoolData and sheetData
    
    :param schoolData: A SchoolData class that contains sample information
                       from each school. Class defined in SchoolDataClass.py
    
    :param sheetData: A SheetData class that contains sample information
                      from the report sheet. Class defined in SheetData.py
                      
    :returns : A 2D array that contains the parent sample name, and the
               row number that in the sheetData excel sheet
    """
    toReturn = [];
    for each_parent_name in schoolData.data.keys():
        current_data = schoolData.data[each_parent_name];
        
        
        for i in range(len(sheetData.quality)):
            compare_result = compare_sample(current_data, i, sheetData);
            
            if (compare_result == True):
                toReturn.append([sheetData.sheet_index[i], each_parent_name, schoolData.data[each_parent_name].getAllImages()])
                break;
            
            
    return toReturn
    
def compare_sample(school_sample, sheet_sample_index, all_sheet_sample):
    
    """ Compares a single data entry in row i from the all_sheet_sample with 
    School_sample.
    :param school_sample: a SampleData class that contain metrology and metallography
                          information for a single entry.
    :param sheet_sample_index: the row number in the all_sheet_sample that the
                               user would like to check against.
    :param all_sheet_sample: a SheetData class variables that contains all the information
                             from the excel sheet
                             
    :returns : True if the school sample matches the sample in the sheet_sample_index
               entry in the all_sheet_sample class, and false otherwise
    
    """
    
    # compare metrology
    school_quality = school_sample.metrology.quality;
    school_bead_width = school_sample.metrology.bead_width;
    
    sheet_quality = all_sheet_sample.quality[sheet_sample_index]
    sheet_bead_width = all_sheet_sample.bead_width[sheet_sample_index]
    
    if (school_quality != sheet_quality) or (school_bead_width != sheet_bead_width):
        return False;
    
    # compare metalography
    school_fz_50 = school_sample.metallography.fz_50;
    school_fz_75 = school_sample.metallography.fz_75;
    
    sheet_fz_50 = all_sheet_sample.fz_50[sheet_sample_index];
    sheet_fz_75 = all_sheet_sample.fz_75[sheet_sample_index];
    
    if (school_fz_50 != sheet_fz_50) or (school_fz_75 != sheet_fz_75):
        return False;
    
    
    return True
  
def getNameIndex(workbook):
    """ Gets the column number corresponding to each variables name in the excel sheet
    
    :param workbook: A xlrd workbook variables that the user want to find the 
                     variable index
                     
    :returns : A tuple consists of Metrollogy_Nameing class and Metallography_Nameing
               class the contains the indexing information. Those two class are in SheetData.py
    """
    name_row = 0;
    metrology_sheet = workbook.sheet_by_name("metrology");
    metallography_sheet = workbook.sheet_by_name("metallography")
    metrology_index = Metrology_Naming(list(metrology_sheet.row_values(name_row)));
    metallography_index = Metallography_Naming(list(metallography_sheet.row_values(name_row))); 
    
    return metrology_index, metallography_index


def extractData(school_name):
    
    """ Extract the metrollogy and metallography information from the corresponding excel
    sheet of the school name.
    
    :param school_name: a string of the school name
    :returns: a dictionary with the key of parent sample names, and elements of SampleData class
              variable
    """
    
    #############################################################
    # opens file and read the metrology and metallography info
    school_name = school_name.lower();
    fileToOpen = "";
    
    if (school_name == 'csm') or (school_name == 'mines'):
        fileToOpen = "Characterization-Mines.xlsx"
        
    elif (school_name == 'cmu'):
        fileToOpen = "Characterization-CMU.xlsx"
        
    workbook = xlrd.open_workbook(fileToOpen);
    metrology = workbook.sheet_by_name("metrology")
    metallography = workbook.sheet_by_name("metallography")
    
    metrology_variables, metallography_variables = getNameIndex(workbook);
    
    # populate the metrology entry for the school
    counter = 0
    metrology_dict = {};
    
    for parent_sample_name in metrology.col_values(metrology_variables.parent_sample_name):
        
        # extract parent name information for the samples
        if not hasNumber(str(parent_sample_name)):
            counter += 1;
            continue;
        
        
        # the metrology information and its related images
        curr_bead_quality = metrology.cell_value(counter, metrology_variables.bead_quality)
        curr_bead_width = metrology.cell_value(counter, metrology_variables.bead_width)
        curr_image_files = metrology.cell_value(counter, metrology_variables.file_location)
        
        # discard the current row if it is empty
        if (not curr_bead_quality) or (not curr_bead_width):
            counter += 1
            continue
        
        
        curr_metrology = Metrology(curr_bead_quality, curr_bead_width, curr_image_files)
        metrology_dict[parent_sample_name] = curr_metrology
            
        counter += 1
    
    
    # populated the metellography entry for the school
    counter = 0;
    metallography_dict = {};
    for parent_sample_name in metallography.col_values(metallography_variables.parent_sample_name):
            
        # extract parent name information for the samples
        if not hasNumber(str(parent_sample_name)):
            counter += 1;
            continue;
            
        # the metallography information and its related images
        curr_fz_area = metallography.cell_value(counter, metallography_variables.fz_area);
        curr_fz_width = metallography.cell_value(counter, metallography_variables.fz_width);
        curr_fz_50 = metallography.cell_value(counter, metallography_variables.fz_50);
        curr_fz_75 = metallography.cell_value(counter, metallography_variables.fz_75);
        curr_image_files = metallography.cell_value(counter, metallography_variables.file_location)
            
            
        # discard the current row if it is empty
        if (not curr_fz_area) or (not curr_fz_width) or (not curr_fz_50) or (not curr_fz_75):
            counter += 1;
            continue
            
        
        curr_metallography = Metallography(curr_fz_width, curr_fz_area, 
                                           curr_fz_50, curr_fz_75, curr_image_files);
                                                   
        metallography_dict[parent_sample_name] = curr_metallography
        counter += 1
        
    toReturn_dict = create_comp_dict(metrology_dict, metallography_dict)     
    return toReturn_dict;
    
if __name__ == '__main__':
    
    ######################################################################
    # This is the data extraction for the report
    ######################################################################
    
   

    report_wb = xlrd.open_workbook(reportPath)
    report_main_sheet = report_wb.sheet_by_name("original")
    
    counter = 0;
    sheetData = SheetData();
    for item in report_main_sheet.col_values(0):
        
        if not item:
            counter += 1;
            continue;
        
        curr_quality = report_main_sheet.cell_value(counter, 4);
        curr_width = report_main_sheet.cell_value(counter, 5);
        curr_fz_area = report_main_sheet.cell_value(counter, 8)
        curr_fz_width = report_main_sheet.cell_value(counter, 12)
        curr_fz_50 = report_main_sheet.cell_value(counter, 10)
        curr_fz_75 = report_main_sheet.cell_value(counter, 11)
        
        
        
        if (not curr_fz_area) or (not isinstance(curr_fz_area, float)):
            
            counter += 1;
            continue;
        
        sheetData.add_entry(counter, curr_quality, curr_width, curr_fz_area, curr_fz_width, curr_fz_50, curr_fz_75)
        counter += 1
        
    ###########################################################
    # finding matched results
    ########################################################   
    cmu_compelete_dict = extractData("CMU");
    cmu_data = SchoolData("CMU");
    cmu_data.data = cmu_compelete_dict;
    
    mines_compelete_dict = extractData("Mines")
    mines_data = SchoolData("CSM")
    mines_data.data = mines_compelete_dict
    
    
    mines_matched_result = findMatchedResults(mines_data, sheetData)
    cmu_matched_result = findMatchedResults(cmu_data, sheetData)
    
    
    
    ##########################################################
    # Write to the report file
    ##########################################################
    toWriteFile = load_workbook(reportPath)
    toWriteSheet = toWriteFile.get_sheet_by_name("original")
    
    
    
    for currentResult in mines_matched_result:
        toWriteIndex = currentResult[0] + 1;
        cellToWrite = "D" + str(toWriteIndex);
        toWriteSheet[cellToWrite].value = str(currentResult[2]);
 
    toWriteFile.save('testFile.xlsx')
    
    for currentResult in cmu_matched_result:
        toWriteIndex = currentResult[0] + 1;
        cellToWrite = "D" + str(toWriteIndex);
        toWriteSheet[cellToWrite].value = str(currentResult[2]);
 
    toWriteFile.save('testFile.xlsx')
   
    
    
    