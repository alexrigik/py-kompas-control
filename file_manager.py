#connecting to Kompas3D
import os
import re
import pythoncom
import subprocess
from win32com.client import Dispatch, gencache

#connection to Kompa3D API7
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}",0,1,0,0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}",0,1,0).constants
    return module, api, const