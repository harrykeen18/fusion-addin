#Author-
#Description-

import adsk.core, adsk.fusion, adsk.cam, traceback

#import system modules
import os, sys 

#get the path of add-in
my_addin_path = os.path.dirname(os.path.realpath(__file__)) 
#hard coded the absolute path of my add-in
print(my_addin_path)

#add the path to the searchable path collection
if not my_addin_path in sys.path:
  sys.path.append(my_addin_path)

from xlrd import open_workbook


FEATURE_DICT = {}
LAYER_DICT = import_prn()

def run(context):
  
  try:

    app = adsk.core.Application.get()
    
    #ui = app.userInterface
    
    design = app.activeProduct

    #get the root component of the active design.
    rootComp = design.rootComponent
    
    #find face features
    for feature in rootComp.features:

      all_faces = feature.faces
      #print(all_faces.count)
      
      faces_areas = []
      
      for face in all_faces:
        faces_areas.append(face.area)
        #print(face.area)

      #find two largest face areas - top and bottom face
      #if both are the same (part with no features)
      #back face is the lower one

      biggest_area = max(faces_areas)
      biggest_area_index = faces_areas.index(biggest_area)
      #print(biggest_area_index)

      next_biggest_area = max(n for n in faces_areas if n != max(faces_areas))
      next_biggest_area_index = faces_areas.index(next_biggest_area)
      #print(next_biggest_area_index)

      if biggest_area == next_biggest_area:
        #part with no features on it - take bottom face as one with lowest z value (not great solution as stuff could be orientated in a different place.)

        biggest_area_point = all_faces[biggest_area_index].pointOnFace().asArray
        next_biggest_area_point = all_faces[next_biggest_area_index].pointOnFace().asArray

        if biggest_area_point[2] < next_biggest_area_point[2]:
          back_face_ind = biggest_area_index
          top_face_ind = next_biggest_area_index
        else:
          back_face_ind = next_biggest_area_index
          top_face_ind = biggest_area
      else:
        back_face_ind = biggest_area_index
        top_face_ind = next_biggest_area_index

          
      #print(str(largest_area))
      #print(str(back_face_ind))
      
      back_face = all_faces[back_face_ind]
      top_face = all_faces[top_face_ind]
      #print(str(back_face.loops.count))

      for loop in back_face.loops:
        if loop.isOuter:
          #out_profile = loop.edges
          print('outer profile')
        
        #inner loops
        else:
          for edge in loop.edges:
            if edge.geometry.curveType == 2:
              #THRU HOLE
              print('thru hole')
              break
            else:
              #INSIDE THRU CUT
              print('thru inside cut')
              break
          
      print('--------------------------------')

      #top and bottom faces index
      exclude_faces = [biggest_area_index, next_biggest_area_index]
      
      for i, face in enumerate(all_faces):
        if face.geometry.surfaceType == 0: #planar surface
          if not face.geometry.normal.isEqualTo(top_face.geometry.normal):
            exclude_faces.append(i)
        else:
          exclude_faces.append(i)
      #print(exclude_faces)

      #get point for referecne on top face
      top_face_point = top_face.pointOnFace
      
      #iterate through all faces that arn't exluded
      for index, face in enumerate(all_faces):
        if index not in exclude_faces:
          #print(index)
          for loop in face.loops:
            if loop.isOuter:
              #out_profile = loop.edges

              for edge in loop.edges:
                if edge.geometry.curveType == 2:
                  #NON THRU HOLE
                  #find distance between faces via dot product.

                  depth = get_depth(face, top_face_point)
                  #print(depth)
                  print('hole ' + str(depth))



                else:
                  #INSIDE NON THRU CUT
                  print('inside cut')
                  break
            
            #inner loops
            else:
              print('some internal feature') #don't need to worry about this.
  except:
    if ui:
      print('Failed:\n{}'.format(traceback.format_exc()))


#def write_dxf(): #give all the dictionaries.


#def layer_name():


def import_prn():
  file_loc = '/Users/harry/Dropbox (OpenDesk)/06_Production/06_Software/CADLine Plugin/excel files/LAYERCOLOURS - new.xlsx'
  wb = open_workbook(file_loc)

  sheet = wb.sheet_by_index(0)

  #row = sheet.row(4)

  sheetdict = {}
  for colnum in range(1, sheet.ncols):
    col_values_list = []
    for rownum in range(1, sheet.nrows):  
      #TO DO loop through each row and append in to a []
      col_values_list.append(eval(sheet.cell_value(rownum, colnum)))
    
    #print(col_values_list)
    sheetdict[sheet.cell_value(0, colnum)] = col_values_list
  #print(sheetdict)

  return sheetdict



def get_depth(face, top_face_point):
  point = face.pointOnFace
  vector = point.vectorTo(top_face_point)
  depth = face.geometry.normal.dotProduct(vector)

  return depth

def start_section(dxf):
  dxf.append('  0')
  dxf.append('SECTION')

  return dxf

def end_section(dxf):
  dxf.append('  0')
  dxf.append('ENDSEC')

  return dxf


def end_dxf(dxf):
  dxf.append("  0")
  dxf.append("SEQEND")

  return dxf


def start_layer(dxf):
  dxf.append('  2')
  dxf.append('TABLE')
  dxf.append('  0')
  dxf.append('TABLE')
  dxf.append('  2')
  dxf.append('LAYER')
  dxf.append(' 70')
  dxf.append('     4')
  dxf.append('  0')
  dxf.append('LAYER')
  dxf.append('  2')
  dxf.append('0')
  dxf.append(' 70')
  dxf.append('     0')
  dxf.append(' 62')
  dxf.append('     7')
  dxf.append('  6')
  dxf.append('CONTINUOUS')

  return dxf


def add_layer(dxf, layer, colour):
  dxf.append('  0')
  dxf.append('LAYER')
  dxf.append('  2')
  dxf.append(layer)
  dxf.append(' 70')
  dxf.append('     0')
  dxf.append(' 62')
  dxf.append('   ' + colour)
  dxf.append('  6')
  dxf.append('CONTINUOUS')

  return dxf


def end_layer(dxf):
  dxf.append('  0')
  dxf.append('ENDTAB')

  return dxf


def start_entities(dxf):
  dxf.append('  2')
  dxf.append('ENTITIES')

  return dxf


def start_polyline(dxf, layer, colour):
  dxf.append('  0')
  dxf.append('POLYLINE')
  dxf.append('8')
  dxf.append(layer)
  dxf.append(' 66')
  dxf.append(colour)
  dxf.append(' 10')
  dxf.append('0.0')
  dxf.append(' 20')
  dxf.append('0.0')
  dxf.append(' 30')
  dxf.append('0.0')

  return dxf


def add_vertex(dxf, layer, point, bulge=0.0):
  dxf.append('  0')
  dxf.append('VERTEX')
  dxf.append('  8')
  dxf.append(layer)
  dxf.append(' 10')
  dxf.append(str(point[0]))
  dxf.append(' 20')
  dxf.append(str(point[1]))
  dxf.append(' 30')
  dxf.append(str(point[2]))

  return dxf


def end_polyline(dxf):
  dxf.append("  0")
  dxf.append("SEQEND")

  return dxf


def circle(dxf, diameter, point, layer):

  dxf.append("  0")
  dxf.append("CIRCLE")
  dxf.append("  8")
  dxf.append(layer)
  dxf.append(' 10')
  dxf.append(str(point[0]))
  dxf.append(' 20')
  dxf.append(str(point[1]))
  dxf.append(' 30')
  dxf.append(str(point[2]))
  dxf.append(" 40")
  dxf.append(diameter)

  return dxf










