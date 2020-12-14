# connecting to Kompas3D
import os
import re
import pythoncom
import subprocess
import yaml
from win32com.client import Dispatch, gencache


# connection to Kompas3D API7
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)  # import KompasAPI7 module
    print("2.1")
    pythoncom.CoInitializeEx(0)  # necessary for pythoncom in django !!!
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    print("2.2")
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1,
                                  0).constants  # import Kompas constants
    print("2.3")
    return module, api, const


def get_kompas_api5():
    module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)  # import KompasAPI5 module
    pythoncom.CoInitializeEx(0)  # necessary for pythoncom in django !!!
    api = module.KompasObject(
        Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(module.KompasObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1,
                                  0).constants  # import Kompas 3D constants
    return module, api, const


# module7, api7, const7 = get_kompas_api7()
# module5, api5, const5 = get_kompas_api5()
#
# app7 = api7.Application  # managing Kompas app

### some feature to test app7 functionality ###
# app7.Visible = True  # show Kompas window
# app7.HideMessage = const7.ksHideMessageNo  # close all Kompas notification
# print(app7.ApplicationName(FullName=True))  # print Kompas program name


def create_file(api, asm, dtl):  # Create file
    doc3D = api.Document3D()
    doc3D.Create(asm, dtl)  # create detail
    return doc3D


def set_obj_type(const, o_type):
    if o_type == "Main_component":
        obj_type = const.pTop_Part
    elif o_type == "Sketch":
        obj_type = const.o3d_sketch
    elif o_type == "Surface":
        obj_type = const.o3d_face
    elif o_type == "Plane_XOY":
        obj_type = const.o3d_planeXOY
    elif o_type == "Plane_XOZ":
        obj_type = const.o3d_planeXOZ
    elif o_type == "Plane_YOZ":
        obj_type = const.o3d_planeYOZ
    elif o_type == "Extrusion":
        obj_type = const.o3d_bossExtrusion
    elif o_type == "Rotation":
        obj_type = const.o3d_bossRotated
    elif o_type == "Direction_forward":
        obj_type = const.dtNormal
    elif o_type == "Direction_reverse":
        obj_type = const.dtReverse
    elif o_type == "Direction_both":
        obj_type = const.dtBoth
    elif o_type == "Direction_middle_plane":
        obj_type = const.dtMiddlePlane
    elif o_type == "Extrusion_str_to_depth":
        obj_type = const.etBlind  # strictly to the depth
    return obj_type


def create_sketch(doc, const, set_plane):
    Part = doc.GetPart(set_obj_type(const, "Main_component"))  # choose main component
    sketch = Part.NewEntity(set_obj_type(const, "Sketch"))
    sketch_definition = sketch.GetDefinition()
    Plane = Part.GetDefaultEntity(set_obj_type(const, set_plane))
    sketch_definition.SetPlane(Plane)
    sketch.Create()
    return sketch, sketch_definition


def edit_sketch(sketch, sketch_def, operations):
    sketch2D = sketch_def.BeginEdit()  ##
    for step in operations:
        if list(step.keys())[0] == "ArcByPoint":
            xc = step["ArcByPoint"]["xc"]
            yc = step["ArcByPoint"]["yc"]
            rad = step["ArcByPoint"]["rad"]
            x1 = step["ArcByPoint"]["x1"]
            y1 = step["ArcByPoint"]["y1"]
            x2 = step["ArcByPoint"]["x2"]
            y2 = step["ArcByPoint"]["y2"]
            direction = step["ArcByPoint"]["direction"]
            style = step["ArcByPoint"]["style"]
            sketch2D.ksArcByPoint(xc, yc, rad, x1, y1, x2, y2, direction, style)
        if list(step.keys())[0] == "ArcBy3Points":
            x1 = step["ArcBy3Points"]["x1"]
            y1 = step["ArcBy3Points"]["y1"]
            x2 = step["ArcBy3Points"]["x2"]
            y2 = step["ArcBy3Points"]["y2"]
            x3 = step["ArcBy3Points"]["x3"]
            y3 = step["ArcBy3Points"]["y3"]
            style = step["ArcBy3Points"]["style"]
            sketch2D.ksArcBy3Points(x1, y1, x2, y2, x3, y3, style)
        elif list(step.keys())[0] == "Circle":
            x = step["Circle"]["x"]
            y = step["Circle"]["y"]
            rad = step["Circle"]["rad"]
            style = step["Circle"]["style"]
            sketch2D.ksCircle(x, y, rad, style)
        # elif list(step.keys())[0] == "Rectangle":
        #     xc = step["Rectangle"]["xc"]
        #     yc = step["Rectangle"]["yc"]
        #     rad = step["Rectangle"]["rad"]
        #     x1 = step["Rectangle"]["x1"]
        #     y1 = step["Rectangle"]["y1"]
        #     x2 = step["Rectangle"]["x2"]
        #     y2 = step["Rectangle"]["y2"]
        #     style = step["Rectangle"]["style"]
        #
        #     definition = obj.GetParamStruct()
        #     ExtrusionParam = definition.ExtrusionParam()
        #     centre = step["Rectangle"]["centre"]
        #     sketch2D.ksRectangle(xc, yc, rad, x1, y1, x2, y2, centre, style)
        elif list(step.keys())[0] == "LineSeg":
            x1 = step["LineSeg"]["x1"]
            y1 = step["LineSeg"]["y1"]
            x2 = step["LineSeg"]["x2"]
            y2 = step["LineSeg"]["y2"]
            style = step["LineSeg"]["style"]
            sketch2D.ksLineSeg(x1, y1, x2, y2, style)
    # sketch2D.ksCircle(0.0,0.0,50.0,1)
    # sketch2D.ksCircle(0.0,0.0,100.0,1)
    sketch_def.EndEdit()  ##


def SetExtrusionParam(defin, const_s,
                      dN=0, dR=0, dirc="Direction_forward",
                      dON=False, dOR=False,
                      dVN=0.0, dVR=0.0,
                      tN="Extrusion_str_to_depth", tR="Extrusion_str_to_depth"
                      ):
    ExtrusionParam = defin.ExtrusionParam()
    ExtrusionParam.depthNormal = dN
    ExtrusionParam.depthReverse = dR
    ExtrusionParam.direction = set_obj_type(const_s, dirc)
    ExtrusionParam.draftOutwardNormal = dON
    ExtrusionParam.draftOutwardReverse = dOR
    ExtrusionParam.draftValueNormal = dVN
    ExtrusionParam.draftValueReverse = dVR
    ExtrusionParam.typeNormal = set_obj_type(const_s, tN)
    ExtrusionParam.typeReverse = set_obj_type(const_s, tR)


def SetColorParam(curr_obj):
    ColorParam = curr_obj.ColorParam()
    ColorParam.ambient = 0.5
    ColorParam.color = 9474192
    ColorParam.diffuse = 0.6
    ColorParam.emission = 0.5
    ColorParam.shininess = 0.8
    ColorParam.specularity = 0.8
    ColorParam.transparency = 1.0


def extrusion(doc, const, sketch, obj_name, dhn, dhr):
    part = doc.GetPart(set_obj_type(const, "Main_component"))
    obj = part.NewEntity(set_obj_type(const, "Extrusion"))

    definition = obj.GetDefinition()
    definition.SetSketch(sketch)

    ## for shim to get edge
    # Collection = Part.EntityCollection(const5.o3d_edge)
    # Collection.SelectByPoint(-100.0,0.0,0.0)
    # Edge = Collection.Last()
    # EdgeDefinition = Edge.GetDefinition()
    # Sketch = EdgeDefinition.GetOwnerEntity()
    ##
    # Extrusion Parameters
    SetExtrusionParam(definition, const, dN=dhn, dR=dhr)
    # ExtrusionParam = definition.ExtrusionParam()
    # ExtrusionParam.depthNormal = 10.0
    # ExtrusionParam.depthReverse = 10.0
    # ExtrusionParam.direction = set_obj_type(const, "Direction_forward")
    # ExtrusionParam.draftOutwardNormal = False
    # ExtrusionParam.draftOutwardReverse = False
    # ExtrusionParam.draftValueNormal = 0.0
    # ExtrusionParam.draftValueReverse = 0.0
    # ExtrusionParam.typeNormal = set_obj_type(const, "Extrusion_str_to_depth")
    # ExtrusionParam.typeReverse = set_obj_type(const, "Extrusion_str_to_depth")

    ThinParam = definition.ThinParam()
    ThinParam.thin = False

    obj.name = obj_name

    SetColorParam(obj)
    # ColorParam = obj.ColorParam()
    # ColorParam.ambient = 0.5
    # ColorParam.color = 9474192
    # ColorParam.diffuse = 0.6
    # ColorParam.emission = 0.5
    # ColorParam.shininess = 0.8
    # ColorParam.specularity = 0.8
    # ColorParam.transparency = 1.0

    obj.Create()  # executing of operation


def SetRotatedParam(defin):
    RotatedParam = defin.RotatedParam()
    RotatedParam.angleNormal = 360.0
    RotatedParam.angleReverse = 0.0
    RotatedParam.direction = 0
    RotatedParam.toroidShape = False

def rotation(doc, const, sketch, obj_name
             # sketch, obj_name, dhn, dhr

             ):
    part = doc.GetPart(set_obj_type(const, "Main_component"))
    obj = part.NewEntity(set_obj_type(const, "Rotation"))
    definition = obj.GetDefinition()
    # print(dir(obj.GetDefinition()))
    definition.SetSketch(sketch)
    # Rotation parameters
    SetRotatedParam(definition)
    # RotatedParam = definition.RotatedParam()
    # print(dir(RotatedParam))
    # SetExtrusionParam(definition, const, dN=dhn, dR=dhr)
    ThinParam = definition.ThinParam()
    ThinParam.thin = False
    obj.name = obj_name
    SetColorParam(obj)
    obj.Create()  # executing of operation

def save_as_and_quite(doc, api, file_location):
    doc.SaveAs(file_location)  # save detail
    api.Quit()  # close Kompas


# creation of washer
# doc3D = create_file(api5, False, True)  # create detail
# Sketch, sketch_definition = create_sketch(doc3D, const5, "Plane_XOY")  # create sketch on XOY plane
# operations = [{"Circle": {"x": 0.0, "y":  0.0, "rad":  50.0, "style":  1}},  # list of 2D operations
#               {"Circle": {"x": 0.0, "y":  0.0, "rad":  100.0, "style":  1}},
#               {"ArcByPoint": {"xc": 150.0, "yc": 150.0, "rad": 20 ,
#                               "x1": 130.0, "y1": 150.0, "x2": 150.0, "y2": 170.0, "direction": -1, "style": 1}},
#               {"LineSeg": {"x1": 130.0, "y1": 150.0, "x2": 150.0, "y2": 170.0, "style": 1}}]
# edit_sketch(Sketch, sketch_definition, operations)  # drawing circle by operations from operations list
# extrusion(doc3D, const5, Sketch, "Extrusion operation 1")  # extrude of washer's body
# save_as_and_quite(doc3D, app7, "D:\/washer.m3d")  # save detail to the file with name washer.m3d


# print(parsed_config_file)

def interpreter(config_file):  # , const5 = const5
    global doc3D
    global plane
    global operations
    global Sketch
    global sketch_definition

    parsed_config_file = yaml.load(config_file, Loader=yaml.FullLoader)
    ####
    module7, api7, const7 = get_kompas_api7()
    print('2')
    module5, api5, const5 = get_kompas_api5()
    print('3')
    app7 = api7.Application  # managing Kompas app
    for i in parsed_config_file["command"]:
        if list(i.keys())[0] == "create":
            if i["create"] == "detail":
                doc3D = create_file(api5, False, True)  # create detail
            elif i["create"] == "assembly":
                doc3D = create_file(api5, True, False)  # create assembly
            elif i["create"] == "sketch":
                Sketch, sketch_definition = create_sketch(doc3D, const5, plane)  # create sketch on XOY plane
            elif list(i["create"].keys())[0] == "extrusion":
                extrusion(doc3D, const5, Sketch,
                          i["create"]["extrusion"]["name"],
                          i["create"]["extrusion"]["depthNormal"],
                          i["create"]["extrusion"]["depthReverse"]
                          )  # extrude of washer's body
            elif list(i["create"].keys())[0] == "rotation":
                rotation(doc3D, const5, Sketch,
                         i["create"]["rotation"]["name"]
                         )
        if list(i.keys())[0] == "edit":
            if i["edit"] == "sketch":
                edit_sketch(Sketch, sketch_definition, operations)  # drawing circle by operations from operations list
        if list(i.keys())[0] == "set_plane":
            if i["set_plane"] == "XOY":
                plane = "Plane_XOY"
            elif i["set_plane"] == "XOZ":
                plane = "Plane_XOZ"
            elif i["set_plane"] == "YOZ":
                plane = "Plane_YOZ"
        if list(i.keys())[0] == "saveAs_quit":
            save_as_and_quite(doc3D, app7,
                              i["saveAs_quit"]["path"] +
                              i["saveAs_quit"]["name"])  # save detail to the file with name washer.m3d
        if list(i.keys())[0] == "draw":
            operations = []
            for j in i["draw"]:
                operations.append(j)


# def yaml_parser(file):
#     #print("+")
#     #yaml_file = open(file)  # loading configuration file
#     parsed_config_file = yaml.load(file, Loader=yaml.FullLoader)  # parsing it
#     print(parsed_config_file)
#     return parsed_config_file

# yaml_file = open("config.yaml")  # loading configuration file
# parsed_config_file = yaml.load(yaml_file, Loader=yaml.FullLoader)  # parsing it
#
# interpreter(parsed_config_file)  ## run interpretor

# some drafts
# doc3D.fileName = "D:\/test.m3d"
# doc3D.SaveAs()
# CurrentDocument = api7.ActiveDocument

# print(dir(const5))
# print(dir(app7))
# print(dir(const7))

# sketch2D.ksCircle(0.0,0.0,50.0,1)
# sketch2D.ksCircle(0.0,0.0,100.0,1)
