"""
@FileName: basic/__init__.py
@Description: 该文件主要提供桌面应用的各种基础操作,包括文件和设计的打开、关闭、删除,以及变量的各种相关操作
@Author: SJY
@Time: 2024/2/7 9:50
@Coding: UTF-8
@Vision: V1.0
"""

# Official Standard Library
import os

# Third-party Library


# Local Library
# import hfss

# Import the OSL with "from"


# Import the TPL with "from"


# Import the LL with "from"
# 引用 hfss 对象
from hfss import oDesktop, oAnsoftApp, oProject, oDesign


# 函数定义
def OpenAnsysElectronicsDesktop():
    """
    该函数用于打开 Ansys 电子桌面
    :return: None
    """
    oDesktop.handle = oAnsoftApp.oDesktop


def CloseAnsysElectronicsDesktop():
    """
    该函数用于关闭 Ansys 电子桌面
    :return: None
    """
    oProject.handle = oDesktop.GetActiveProject()
    if oProject.handle is not None:
        oDesktop.CloseProject(oProject.GetName())
    oDesktop.QuitApplication()
    # del oProject, oDesign, oDesktop, oAnsoftApp


def OpenProject(project_path: str):
    """
    该函数用于打开一个已存在的工程文件
    :param project_path: 工程文件路径
    :return None:
    """
    try:
        if not os.path.exists(project_path):
            raise FileNotFoundError
        # 分离文件路径和文件名
        dir_path, file_name_ext = os.path.split(project_path)
        file_name, file_ext = os.path.splitext(file_name_ext)
        if file_ext != ".aedt":
            raise TypeError
        oProject.handle = oDesktop.GetActiveProject()
        if oProject.handle is not None:
            if oProject.GetName() == file_name:
                # print("Project already open")
                pass
            else:
                oDesktop.OpenProject(project_path)
        else:
            oDesktop.OpenProject(project_path)
    except FileNotFoundError:
        print("FileNotFoundError")
    except TypeError:
        print("TypeError")
    pass


def CreateProject(project_name):
    """
    创建一个新工程
    :param project_name: 得包含路径名称 新工程名字
    :return:
    """
    oDesktop.NewProject()
    oProject.handle = oDesktop.GetActiveProject()
    oProject.Rename(project_name, False)

    pass


def CloseProject(project_name):
    """
    在桌面中关闭项目,但不退出应用
    :param project_name:
    :return:
    """
    oProject.handle = oDesktop.GetActiveProject()
    if oProject.handle is not None:
        oProject.Save()
    project_name_list = oDesktop.GetProjectList()
    if project_name in project_name_list:
        oDesktop.CloseProject(project_name)
    else:
        print(project_name, "is non-existence")
    pass

def CopyHFSSDesign(design_name, new_design_name):
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
        pass
    design_name_list = oProject.GetTopDesignList()
    if design_name in design_name_list:
        oProject.CopyDesign(design_name)
    else:
        print(design_name, "is non-existence.")
        return 0
    if new_design_name in design_name_list:
        print(new_design_name, "is already existing.")
        pass
    else:
        oProject.Paste()
        oDesign.handle = oProject.GetActiveDesign()
        if oDesign.handle is None:
            # design_name_list = oProject.GetTopDesignList()
            # oDesign.handle = oProject.GetDesign(design_name_list[-1])
            oDesign.handle = oProject.GetDesign("HFSSDesign13")
            pass
        name = oDesign.GetName()
        oDesign.RenameDesignInstance(name, new_design_name)
    oProject.Save()
    pass

def DeleteHFSSDesign(design_name):
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
        pass
    design_name_list = oProject.GetTopDesignList()
    if design_name in design_name_list:
        oProject.DeleteDesign(design_name)
    else:
        print(design_name,"is non-existence.")
        pass
    oProject.Save()
    pass


def InsertHFSSDesign(design_name, solution_type):
    """
    新建一个设计
    :param design_name: 新建设计名称
    :param solution_type: 求解类型 "DrivenModal", "DrivenTerminal", or "Eigenmode".
    :return:
    """
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    oProject.InsertDesign("HFSS", design_name, solution_type)


def CreateNewVariable(props_name_array, props_value_array):
    """
    在 active 设计中批量创建一些长度变量
    :param props_name_array: 字符串 列表,每个元素为字符串
    :param props_value_array: 浮点数 列表,每个元素为浮点数,单位为 mm
    :return:
    """
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    for i in range(len(props_name_array)):
        oDesign.ChangeProperty(["NAME:AllTabs",
                                ["NAME:LocalVariableTab",
                                 ["NAME:PropServers", "LocalVariables"],
                                 ["NAME:NewProps",
                                  ["NAME:"+props_name_array[i], "PropType:=", "VariableProp", "UserDef:=", True,
                                   "Value:=", str(props_value_array[i])+"mm"]]]])


def GetVariableName() -> tuple:
    """
    得到 active 设计中所有变量的名称
    :return: tuple 包含当前设计所有变量的名称 字符串元组
    """
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    return oDesign.GetVariables()


def GetVariableValue(variable_name):
    """
    得到 active 设计中变量variable_name的值
    :return: 字符串
    """
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    return oDesign.GetVariablesValue(variable_name)
    pass


def ChangeVariable(props_name_array, props_value_array):
    """
    批量修改 active 设计 变量的值
    :param props_name_array: 字符串 列表,每个元素为字符串
    :param props_value_array: 浮点数 列表,每个元素为浮点数,单位为 mm
    :return:
    """
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    # oDesign.handle = oProject.GetActiveDesign()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetDesign("HFSSDesign13")
        print(type(oDesign))
    for i in range(len(props_name_array)):
        oDesign.ChangeProperty(["NAME:AllTabs",
                                ["NAME:LocalVariableTab",
                                 ["NAME:PropServers", "LocalVariables"],
                                 ["NAME:ChangedProps",
                                  ["NAME:"+props_name_array[i], "PropType:=", "VariableProp", "UserDef:=", True,
                                   "Value:=", str(props_value_array[i])+"mm"]]]])


# 函数测试
if __name__ == '__main__':
    projectPath = "P:\\python\\IntelligentAlgorithm\\AntennaData\\ProjectStudy.aedt"
    OpenProject(projectPath)
    # print(oProject.GetTopDesignList())
    # oProject.GetTopDesignList()
    CopyHFSSDesign("HFSSDesign12", "TestDesign")

    # OpenProject("P:\\Ansoft\\myTestProject.aedt")  #  创建新工程
    # InsertHFSSDesign("hfssDesign1", "DrivenModal")
    # CreateNewVariable(["length", "width"],[12, 15])

    # oProject.handle = oDesktop.GetActiveProject()
    # oDesign.handle = oProject.GetActiveDesign()
    # Variable = oDesign.GetVariables()
    # print(Variable, type(Variable))
    # CloseAnsysElectronicsDesktop()  # 关闭应用

    pass
