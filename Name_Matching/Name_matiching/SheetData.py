
class SheetData:
    """ This class contains data from the report sheet
    """
    def __init__(self):
        self.sheet_index = []
        self.quality = []
        self.bead_width = []
        self.fz_area = []
        self.fz_width = []
        self.fz_50 = []
        self.fz_75 = []
        
    def add_entry(self, sheet_index, quality, bead_width, fz_area, fz_width, fz_50, fz_75):
        self.sheet_index.append(sheet_index)
        self.quality.append(quality)
        self.bead_width.append(bead_width)
        self.fz_area.append(fz_area)
        self.fz_width.append(fz_width)
        self.fz_50.append(fz_50)
        self.fz_75.append(fz_75)
        
        
class Metallography_Naming:
    """ This class contains the index for different variable names found in the metallography sheet
    """
    def __init__(self, name_list):
        self.name_list = name_list;
        
        
        self.parent_sample_name = -1;
        
        self.fz_area = -1;
        self.fz_width = -1;
        self.fz_depth = -1;
        self.fz_50 = -1;
        self.fz_75 = -1;
        self.file_location = -1;
        self.organize_names()
        
    def organize_names(self):
        
        
        counter = 0;
        for current_name in self.name_list:
            current_name = current_name.lower();
            
            if "parent sample name" in current_name:
                self.parent_sample_name = counter;
            elif "fusion zone depth at 50%" in current_name:
                self.fz_50 = counter;
            elif "fusion zone depth at 75%" in current_name:
                self.fz_75 = counter;
            elif "fusion zone area" in current_name:
                self.fz_area = counter;
            elif "fusion zone width" in current_name:
                self.fz_width = counter;
            elif "fusion zone depth" in current_name:
                self.fz_depth = counter;
            elif "file" in current_name:
                self.file_location = counter;
            
                
                
            counter += 1;
            
            
class Metrology_Naming:
    """This class contains the index for different variable names found in the metrollogy sheet
    """
    def __init__(self, name_list):
        self.name_list = name_list;
        self.parent_sample_name = -1;
        self.bead_quality = -1;
        self.bead_width = -1;
        self.file_location = -1;
        self.organize_name();
        
        
    def organize_name(self):
        counter = 0;
        
        
        for current_name in self.name_list:
            current_name = current_name.lower();
            
            if "parent sample name" in current_name:
                self.parent_sample_name = counter;
            elif "bead quality" in current_name:
                self.bead_quality = counter;
            elif "bead width" in current_name:
                self.bead_width = counter;
            elif "file" in current_name:
                self.file_location = counter;
                
                
            counter += 1;                                                                                                                
        
            
            