from builtins import staticmethod


class SchoolData:
    """
    A class the contains the the metrology and metallography information
    """
    def __init__(self, schoolName):
        """
        :param schoolName: A string contains the school name
        """
        self.schoolName = schoolName
        self.data = {}
        
    def appendNewData(self, parent_name, sample_data):
        keys = list(self.data.keys());
        init_keys = keys.copy();
        keys.append(parent_name);
        
        initKeyLen  = len(set(init_keys))
        newKeyLen = len(set(keys))
        
        
        
            
        # Check for repeating image name, if exists, return false
        if(newKeyLen == initKeyLen):
            return False
        
        # if this is a new image, add that to the dictionary
        self.data[parent_name] = sample_data
        return True
        
        
        
class SampleData:
    """
    This class contains Metrology and Metallography info
    """
    def __init__(self, metrology, metallography):
        self.metrology = metrology 
        self.metallography = metallography
        
        
    def addMetrology(self, metrology):
        self.metrology = metrology

    def addMetallography(self, metallography):
        self.metallography = metallography
        
    def getAllImages(self):
        return self.metallography.file_names + self.metrology.file_names
    
    
    
    @staticmethod
    def splitImageFiles(file_names):
        file_names = file_names.replace("[", "");
        file_names = file_names.replace("]", "");
        file_names = file_names.split(", ")
        return file_names
    
        
class Metallography:
    """This class contains measurement and image files from Metallography sheet
    """
    def __init__(self, fz_area, fz_width, fz_50, fz_75, file_names):
        
        self.fz_area = fz_area
        self.fz_width = fz_width
        self.fz_50 = fz_50
        self.fz_75 = fz_75
        self.file_names = SampleData.splitImageFiles(file_names)
        
        
class Metrology:
    """This class contains measurement and image files from Metrollogy sheet
    """
    def __init__(self, quality, bead_width, file_names):   
        self.quality = quality
        self.bead_width = bead_width
        self.file_names = SampleData.splitImageFiles(file_names)     
        
        
        