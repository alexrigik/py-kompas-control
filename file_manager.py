# connecting to Kompas3D
import os
import re
import pythoncom
import subprocess
from win32com.client import Dispatch, gencache


# connection to Kompas3D API7
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)  # import KompasAPI7 module
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants  # import Kompas constants
    return module, api, const

def get_kompas_api5():
    module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)  # import KompasAPI5 module
    api = module.KompasObject(
        Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(module.KompasObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1,
                                  0).constants  # import Kompas 3D constants
    return module, api, const

module7, api7, const7 = get_kompas_api7()
module5, api5, const5 = get_kompas_api5()

app7 = api7.Application  # managing Kompas app

### some feature to test app7 functionality ###
# app7.Visible = True  # show Kompas window
# app7.HideMessage = const7.ksHideMessageNo  # close all Kompas notification
# print(app7.ApplicationName(FullName=True))  # print Kompas program name


def create_file(api, asm, dtl):  # Create file
    doc3D = api.Document3D()
    doc3D.Create(asm, dtl)  # create detail
    return doc3D


def set_obj_type(doc, const, type):
        if type == "Main_component":
            obj_type = const5.pTop_Part
        elif type == "Sketch":
            obj_type = const5.o3d_sketch
        elif type == "Surface":
            obj_type = const5.o3d_face
        elif type == "Plane_XOY":
            obj_type = const5.o3d_planeXOY
        elif type == "Plane_XOZ":
            obj_type = const5.o3d_planeXOZ
        elif type == "Plane_YOZ":
            obj_type = const5.o3d_planeYOZ
        elif type == "Extrusion":
            obj_type = const5.o3d_bossExtrusion
        elif type == "Direction_forward":
            obj_type = const5.dtNormal
        elif type == "Direction_reverse":
            obj_type = const5.dtReverse
        elif type == "Direction_both":
            obj_type = const5.dtBoth
        elif type == "Direction_middle_plane":
            obj_type = const5.dtMiddlePlane
        elif type == "Extrusion_str_to_depth":
            obj_type = const5.etBlind  # strictly to the depth
        return obj_type


def create_sketch(doc, const, set_plane):
    Part = doc.GetPart(set_obj_type(doc3D, const5, "Main_component"))  # choose main component
    sketch = Part.NewEntity(set_obj_type(doc3D, const5, "Sketch"))
    sketch_definition = sketch.GetDefinition()
    Plane = Part.GetDefaultEntity(set_obj_type(doc3D, const5, set_plane))
    sketch_definition.SetPlane(Plane)
    sketch.Create()
    return sketch, sketch_definition


def edit_sketch(sketch, sketch_definition, operations):
    sketch2D = sketch_definition.BeginEdit() ##
    for step in operations:
        if list(step.keys())[0] == "Circle":
            x = step["Circle"]["x"]
            y = step["Circle"]["y"]
            rad = step["Circle"]["rad"]
            style = step["Circle"]["style"]
            sketch2D.ksCircle(x, y, rad, style)
    #sketch2D.ksCircle(0.0,0.0,50.0,1)
    #sketch2D.ksCircle(0.0,0.0,100.0,1)
    sketch_definition.EndEdit() ##


def extrusion(doc, const5, sketch, obj_name):
    part = doc.GetPart(set_obj_type(doc3D, const5, "Main_component"))
    obj = part.NewEntity(set_obj_type(doc3D, const5, "Extrusion"))

    definition = obj.GetDefinition()
    definition.SetSketch(sketch)
## for shim to get edge
#Collection = Part.EntityCollection(const5.o3d_edge)
#Collection.SelectByPoint(-100.0,0.0,0.0)
#Edge = Collection.Last()
#EdgeDefinition = Edge.GetDefinition()
#Sketch = EdgeDefinition.GetOwnerEntity()
##
    # Extrusion Parameters
    ExtrusionParam = definition.ExtrusionParam()
    ExtrusionParam.depthNormal = 10.0
    ExtrusionParam.depthReverse = 10.0
    ExtrusionParam.direction = set_obj_type(doc3D, const5, "Direction_forward")
    ExtrusionParam.draftOutwardNormal = False
    ExtrusionParam.draftOutwardReverse = False
    ExtrusionParam.draftValueNormal = 0.0
    ExtrusionParam.draftValueReverse = 0.0
    ExtrusionParam.typeNormal = set_obj_type(doc3D, const5, "Extrusion_str_to_depth")
    ExtrusionParam.typeReverse = set_obj_type(doc3D, const5, "Extrusion_str_to_depth")

    ThinParam = definition.ThinParam()
    ThinParam.thin = False

    obj.name = obj_name

    ColorParam = obj.ColorParam()
    ColorParam.ambient = 0.5
    ColorParam.color = 9474192
    ColorParam.diffuse = 0.6
    ColorParam.emission = 0.5
    ColorParam.shininess = 0.8
    ColorParam.specularity = 0.8
    ColorParam.transparency = 1.0

    obj.Create()  # executing of operation


def save_as_and_quite(doc, api, file_location):
    doc.SaveAs(file_location)  # save detail
    api.Quit()  # close Kompas


# creation of washer
doc3D = create_file(api5, False, True)  # create detail
Sketch, sketch_definition = create_sketch(doc3D, const5, "Plane_XOY")  # create sketch on XOY plane
operations = [{"Circle": {"x": 0.0, "y":  0.0, "rad":  50.0, "style":  1}},  # list of 2D operations
              {"Circle": {"x": 0.0, "y":  0.0, "rad":  100.0, "style":  1}}]
edit_sketch(Sketch, sketch_definition, operations)  # drawing circle by operations from operations list
extrusion(doc3D, const5, Sketch, "Extrusion operation 1")  # extrude of washer's body
save_as_and_quite(doc3D, app7, "D:\/washer.m3d")  # save detail to the file with name washer.m3d



# some drafts
#doc3D.fileName = "D:\/test.m3d"
#doc3D.SaveAs()
#CurrentDocument = api7.ActiveDocument

#print(dir(const5))
#print(dir(app7))
#print(dir(const7))

# sketch2D.ksCircle(0.0,0.0,50.0,1)
# sketch2D.ksCircle(0.0,0.0,100.0,1)
