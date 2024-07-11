#region library
import clr 
import os
import sys
clr.AddReference("System")
import System

clr.AddReference("RevitServices")
import RevitServices
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Collections.Generic import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *


from System.Windows import MessageBox
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Markup import XamlReader
#endregion

#region revit infor
# Get the directory path of the script.py & the Window.xaml
dir_path = os.path.dirname(os.path.realpath(__file__))
#xaml_file_path = os.path.join(dir_path, "Window.xaml")

#Get UIDocument, Document, UIApplication, Application
uidoc = __revit__.ActiveUIDocument
uiapp = UIApplication(uidoc.Document.Application)
app = uiapp.Application
doc = uidoc.Document
activeView = doc.ActiveView

#general infor 
all_FillPatterns = FilteredElementCollector(doc).OfClass(FillPatternElement).WhereElementIsNotElementType().ToElements()
all_grids = FilteredElementCollector(doc).OfClass(Grid).WhereElementIsNotElementType().ToElements()
magenta_color = Color(255, 0, 255)
yellow_color = Color(255, 250, 0)
PATTERN_NAME = "<Solid fill>"

#endregion

#region method
class ColumItem:
    def __init__(self, column, grid_0, grid_1):
        self.Column = column
        self.Grid0 = grid_0
        self.Grid1 = grid_1

class Utils:

    def get_colum_item(self, column):
        original_mark = str(column.get_Parameter(BuiltInParameter.COLUMN_LOCATION_MARK).AsString())
        temp_mark = original_mark

        if original_mark == "": return None
        else:
            #find first grid_0
            grid_0 = ""
            for grid in all_grids:
                if temp_mark.startswith(grid.Name):
                    grid_0 = grid
                    break
            
            #clear first grid_mark
            temp_mark = temp_mark.replace(grid_0.Name,"")
            first_char = temp_mark[0]

            index = 0
            if first_char == '(':
                index = temp_mark.find(')')+2
            else:
                index = temp_mark.find('-')+1

            #get first/second grid_mark
            second_grid_mark = temp_mark[index:]
            first_grid_mark = original_mark.replace(second_grid_mark,"")[:-1]

            #find grid_1
            grid_1 = ""
            for grid in all_grids:
                if second_grid_mark.startswith(grid.Name):
                    grid_1 = grid
                    break
            
            distance_0 = first_grid_mark.replace(grid_0.Name,"")
            distance_1 = second_grid_mark.replace(grid_1.Name,"")

            if distance_0 == "" and distance_1 == "": return ColumItem(column, grid_0, grid_1)
            if distance_0 == "" and distance_1 != "": return ColumItem(column, grid_0, None)
            if distance_0 != "" and distance_1 == "": return ColumItem(column, grid_1, None)
            

            

    def extend_line (self, line, distance):
        if distance == 0: return line
        else:
            sp = line.GetEndPoint(0)
            ep = line.GetEndPoint(1)
            normalize = (ep - sp).Normalize()
            nsp = sp - normalize * distance
            nep = ep + normalize * distance
            
            return Line.CreateBound(nsp, nep)


    def check_columm_and_grid (self, column, grid):
        location = column.Location.Point
        location = XYZ(location.X, location.Y, 0)

        grid_line = None
        if isinstance(grid.Curve, Line):
            point_0 = grid.Curve.GetEndPoint(0)
            point_1 = grid.Curve.GetEndPoint(1)
            
            sp = XYZ(point_0.X, point_0.Y, 0)
            ep = XYZ(point_1.X, point_1.Y, 0)

            line = Line.CreateBound(sp, ep)
            grid_line = self.extend_line(line, 5000)

            distance = grid_line.Distance(location)*304.8
            distance = round(distance, 0)

            if distance == 0: return True

        return False
        
    def highlight_color(self, columnList, color):
        setting = OverrideGraphicSettings()
        setting.SetCutForegroundPatternColor(color)
        setting.SetCutBackgroundPatternColor(color)
        setting.SetSurfaceBackgroundPatternColor(color)
        setting.SetSurfaceForegroundPatternColor(color)

        for pattern in all_FillPatterns:
            if pattern.Name == PATTERN_NAME:
                setting.SetCutBackgroundPatternId(pattern.Id)
                setting.SetCutForegroundPatternId(pattern.Id)
                setting.SetSurfaceBackgroundPatternId(pattern.Id)
                setting.SetSurfaceForegroundPatternId(pattern.Id)
                break

        for column in columnList:
            activeView.SetElementOverrides(column.Id, setting)

    def reset_color(self, columnList):
        setting = OverrideGraphicSettings()
        for column in columnList:
            activeView.SetElementOverrides(column.Id, setting)
            

    def check_colums_location(self, columnsList):
        column_items = []
        columns_to_highlight = []

        for column in columnsList:
            item = self.get_colum_item(column)
            if item is not None:
                column_items.append(item)

        for item in column_items:
            grid0_checked = self.check_columm_and_grid(item.Column, item.Grid0)
            if grid0_checked == False:
                columns_to_highlight.append(item.Column)
            else:
                if item.Grid1 is not None:
                    grid1_checked = self.check_columm_and_grid(item.Column, item.Grid1)
                    if grid1_checked == False: 
                        columns_to_highlight.append(item.Column)
                    
        columns_to_reset = []
        for column in columnsList:
            index = column.get_Parameter(BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM).AsInteger()
            if columns_to_highlight.__contains__(column) == False and index == 0:
                columns_to_reset.append(column)
        
        # highlight and reset color
        self.highlight_color(columns_to_highlight, magenta_color)
        self.reset_color(columns_to_reset)
        

        message = "Found "+ str(len(columns_to_highlight)) + " columns that need to check the position!"
        MessageBox.Show(message, "Message")

#endregion

class FilterColumn(ISelectionFilter):
    def AllowElement(self, element):
        if element.Category.Name == "Structural Columns": return True
        else: return False
             
    def AllowReference(self, reference, position):
        return True

#select elements
class Main ():
    def main_task(self):

        selected_objects = []
        try:
            selected_objects = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, FilterColumn())
        except:
            pass
        
        if len(selected_objects) > 0:

            #classify to vertical and slanted
            vertical_columns = []
            slanted_columns = []
            for r in selected_objects:
                column =doc.GetElement(r)
                index = column.get_Parameter(BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM).AsInteger()
                if index == 0:
                    vertical_columns.append(column)
                else: slanted_columns.append(column)

            #check columns location
            try:
                t = Transaction(doc, " ")
                t.Start()

                Utils().check_colums_location(vertical_columns)
                Utils().highlight_color(slanted_columns, yellow_color)

                t.Commit()
                    
            except Exception as e:
                MessageBox.Show(str(e), "Message")
        

if __name__ == "__main__":
    Main().main_task()
                
    
    





