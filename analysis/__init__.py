"""
@FileName: analysis/__init__.py
@Description: 该文件主要提供关于分析的各种操作功能,设置求解参数、设置远场球坐标、分析、仿真数据导出及处理等功能
@Author: SJY
@Time: 2024/2/7 22:23
@Coding: UTF-8
@Vision: V1.0
更新日志: 目前已完成天线工程的分析仿真、仿真数据(轴比、阻抗带宽)图像的生成,图像在形成的同时,将仿真数据保存至工程路径下的
<工程名>.system_analysis_data 路径下.


"""

# Official Standard Library
import os
# import csv

# Third-party Library
import pandas as pd

# Local Library
# import hfss

# Import the OSL with "from"


# Import the TPL with "from"


# Import the LL with "from"
# 引入 hfss 控制对象
from hfss import oDesktop, oProject, oDesign, oModule


def InsertRadFieldSphereSetup():
    """
    插入远场球坐标 Theta 起始坐标、结束坐标及步进弧度, Phi 起始坐标、结束坐标及步进弧度
    :return:
    """
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    oModule.handle = oDesign.GetModule("RadField")
    oModule.InsertFarFieldSphereSetup(
        [
            "NAME:Infinite Sphere1",
            "UseCustomRadiationSurface:=", False,
            "CSDefinition:=", "Theta-Phi",
            "Polarization:=", "Linear",
            "ThetaStart:=", "-180deg",
            "ThetaStop:=", "180deg",
            "ThetaStep:=", "1deg",
            "PhiStart:=", "0deg",
            "PhiStop:=", "360deg",
            "PhiStep:=", "1deg",
            "UseLocalCS:=", False
        ]
    )


def InsertSetup(setup_name: str, center_freq: float, max_delta: float, setup_type: str,
                start_freq: float, stop_freq: float, range_count: int):
    """
    插入求解设置
    :param setup_name: 求解设置名称
    :param center_freq: 求解中心频率 单位为 GHz
    :param max_delta: 求解最大误差
    :param setup_type: 频点遍历类型
    :param start_freq: 起始频率 单位为 GHz
    :param stop_freq: 结束频率 单位为 GHz
    :param range_count: 求解点数
    :return:
    """
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    oModule.handle = oDesign.GetModule("AnalysisSetup")
    oModule.InsertSetup(
        "HfssDriven",
        [
            "NAME:" + setup_name,
            "SolveType:=", "Single",
            "Frequency:=", str(center_freq) + "GHz",
            "MaxDeltaS:=", max_delta,
            "UseMatrixConv:=", False,
            "MaximumPasses:=", 99, "MinimumPasses:=", 1,
            "MinimumConvergedPasses:=", 3,
            "PercentRefinement:=", 30,
            "IsEnabled:=", True,
            ["NAME:MeshLink", "ImportMesh:=", False],
            "BasisOrder:=", 1,
            "DoLambdaRefine:=", True,
            "DoMaterialLambda:=", True,
            "SetLambdaTarget:=", False,
            "Target:=", 0.3333,
            "UseMaxTetIncrease:=", False,
            "PortAccuracy:=", 2,
            "UseABCOnPort:=", False,
            "SetPortMinMaxTri:=", False,
            "UseDomains:=", False,
            "UseIterativeSolver:=", False,
            "SaveRadFieldsOnly:=", False,
            "SaveAnyFields:=", True,
            "IESolverType:=", "Auto",
            "LambdaTargetForIESolver:=", 0.15,
            "UseDefaultLambdaTgtForIESolver:=", True,
            "IE Solver Accuracy:=", "Balanced"
        ]
    )
    pass

    oModule.InsertFrequencySweep(
        setup_name,
        [
            "NAME:Sweep",
            "IsEnabled:=", True,
            "RangeType:=", "LinearCount",
            "RangeStart:=", str(start_freq) + "GHz",
            "RangeEnd:=", str(stop_freq) + "GHz",
            "RangeCount:=", range_count,
            "Type:=", setup_type,
            "SaveFields:=", True,  # 保存场求解数据
            "SaveRadFields:=", True,  # 保存远场求解数据
            "GenerateFieldsForAllFreqs:=", False
        ]
    )
    pass


def Analyze(setup_name):
    """
    进行仿真分析
    :param setup_name:
    :return: None
    """
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
        pass
    oModule.handle = oDesign.GetModule("AnalysisSetup")
    names = oModule.GetSetups()
    if setup_name in names:
        oDesign.Analyze(setup_name)
    else:
        print("该设计中没有该setup")


def ExportAnalysisDataToFile(file_path: str):
    """
    将仿真数据导出至文件 仿真数据包括 s11参数 阻抗带宽 轴比 轴比带宽 增益
    :param file_path:
    :return: None
    """
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    oModule.handle = oDesign.GetModule("Solutions")
    if file_path is None:
        export_data_file_path = "P://python//hfss-api//data_path"
    else:
        export_data_file_path = file_path
    oModule.handle = oDesign.GetModule("Solutions")
    oModule.ExportNetworkData(
        "",  # Empty string.
        ["Setup1:Sweep"],  # This is the full path of the file from which the solution is loaded.
        3,  # File Type, The number 1.
        export_data_file_path + "NetworkData" + ".s4p",  # full path of file to export to
        ["all"],  # optional, if none defined all frequencies are used
        False, 50,
        "S",
        -1, 0, 15, True, True
    )
    pass


def GenerateS11Graph():
    """
    生成 S11 参数图像 单位为 dB
    生成图表命名为 S Parameter Plot 1
    :return: none
    """
    local_var_array = oDesign.GetVariables()
    local_var_list = ["Freq:=", ["All"]]
    for i in range(len(local_var_array)):
        local_var_list = local_var_list + [local_var_array[i] + ":=", ["Nominal"]]
        pass
    # print(local_var_list)
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    oModule.handle = oDesign.GetModule("ReportSetup")
    ReportNames = oModule.GetAllReportNames()
    # 检索 S11 参数报告是否存在
    if "S Parameter Plot 1" not in ReportNames:
        # 不存在就创建新数据表
        oModule.CreateReport("S Parameter Plot 1", "Modal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
                             ["Domain:=", "Sweep"],
                             local_var_list,
                             [
                                 "X Component:=", "Freq",
                                 "Y Component:=", ["dB(S(1,1))"]
                             ])
    else:  # 存在就仅更新报告
        oModule.UpdateAllReports()
    # print("S Parameter 图表创建成功")
    project_path = oProject.GetPath()  # 获得当前工程文件路径
    project_name = oProject.GetName()  # 获得当前工程名
    data_path = project_path + "/" + project_name + ".system_analysis_data"
    if os.path.exists(data_path) is False:
        os.makedirs(data_path)  # 如果不存在则需创建文件夹
        pass
    # 导出数据文件
    oModule.ExportToFile("S Parameter Plot 1", data_path + "/S Parameter Plot 1.csv", False)
    # print("S Parameter 数据导出成功")
    pass


def GenerateARBWGraph():
    """
    生成轴比带宽图像 theta:= 0 deg, phi:= 0 deg
    生成图表 命名为 Axial Ratio BW Plot 1
    :return: None
    """
    local_var_array = oDesign.GetVariables()
    local_var_list = ["Freq:=", ["All"], "Phi:=", ["0deg"], "Theta:=", ["0deg"]]
    for i in range(len(local_var_array)):
        local_var_list = local_var_list + [local_var_array[i] + ":=", ["Nominal"]]
        pass
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    oModule.handle = oDesign.GetModule("ReportSetup")
    ReportNames = oModule.GetAllReportNames()
    # 检索 AR BW 参数报告是否存在
    if "Axial Ratio BW Plot 1" not in ReportNames:
        # 不存在就创建新数据表
        oModule.CreateReport("Axial Ratio BW Plot 1", "Far Fields", "Rectangular Plot", "Setup1 : Sweep",
                             ["Context:=", "Infinite Sphere1"],
                             local_var_list,
                             [
                                 "X Component:=", "Freq",
                                 "Y Component:=", ["dB(AxialRatioValue)"]
                             ])
    else:  # 存在就仅更新报告
        oModule.UpdateReports("Axial Ratio BW Plot 1")
    # print("Axial Ratio BW 图表创建成功")
    project_path = oProject.GetPath()  # 获得当前工程文件路径
    project_name = oProject.GetName()  # 获得当前工程名
    data_path = project_path + "/" + project_name + ".system_analysis_data"
    if os.path.exists(data_path) is False:
        os.makedirs(data_path)  # 如果不存在则需创建文件夹
        pass
    # 导出数据文件
    oModule.ExportToFile("Axial Ratio BW Plot 1", data_path + "/Axial Ratio BW Plot 1.csv", False)
    # print("Axial Ratio BW  数据导出成功")
    pass


def GenerateRadiationPattern(frequency: float):
    """
    生成指定频率下 2D 辐射图 Gain
    生成图表 命名为 Gain 2D Radiation Pattern Plot 1
    :return: None
    """
    local_var_array = oDesign.GetVariables()
    local_var_list = ["Theta:=", ["All"], "Phi:=", ["0deg", "90deg"], "Freq:=", [str(frequency) + "GHz"]]
    for i in range(len(local_var_array)):
        local_var_list = local_var_list + [local_var_array[i] + ":=", ["Nominal"]]
        pass
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    oModule.handle = oDesign.GetModule("ReportSetup")
    ReportNames = oModule.GetAllReportNames()
    # 检索 Gain 2D Radiation Pattern Plot 1 参数报告是否存在
    if "Gain 2D Radiation Pattern Plot 1" not in ReportNames:
        # 不存在就创建新数据表
        oModule.CreateReport("Gain 2D Radiation Pattern Plot 1", "Far Fields", "Radiation Pattern", "Setup1 : Sweep",
                             ["Context:=", "Infinite Sphere1"],
                             local_var_list,
                             [
                                 "Ang Component:=", "Theta",
                                 "Mag Component:=", ["dB(RealizedGainLHCP)", "dB(RealizedGainRHCP)"]
                             ])
    else:  # 存在就仅更新报告
        oModule.UpdateReports("Gain 2D Radiation Pattern Plot 1")
    # print("Gain 2D Radiation Pattern Plot 1 图表创建成功")
    project_path = oProject.GetPath()  # 获得当前工程文件路径
    project_name = oProject.GetName()  # 获得当前工程名
    data_path = project_path + "/" + project_name + ".system_analysis_data"
    if os.path.exists(data_path) is False:
        os.makedirs(data_path)  # 如果不存在则需创建文件夹
        pass
    # 导出数据文件
    oModule.ExportToFile("Gain 2D Radiation Pattern Plot 1", data_path + "/Gain 2D Radiation Pattern Plot 1.csv", False)
    # print("Gain 2D Radiation Pattern Plot 1  数据导出成功")
    pass


def Generate3DGainRadiationPattern(frequency: float):
    """
    生成 指定频率下 3D 极坐标 辐射模式图
    生成图表命名为 Gain 3D Radiation Pattern Plot 1
    :frequency: 单位 GHz
    :return: None
    """
    local_var_array = oDesign.GetVariables()
    local_var_list = ["Theta:=", ["All"], "Phi:=", ["All"], "Freq:=", [str(frequency) + "GHz"]]
    for i in range(len(local_var_array)):
        local_var_list = local_var_list + [local_var_array[i] + ":=", ["Nominal"]]
        pass
    oModule.handle = oDesign.GetModule("ReportSetup")
    if oProject.handle is None:
        oProject.handle = oDesktop.GetActiveProject()
    if oDesign.handle is None:
        oDesign.handle = oProject.GetActiveDesign()
    oModule.handle = oDesign.GetModule("ReportSetup")
    ReportNames = oModule.GetAllReportNames()
    if "Gain 3D Radiation Pattern Plot 1" not in ReportNames:
        oModule.CreateReport("Gain 3D Radiation Pattern Plot 1", "Far Fields", "3D Polar Plot", "Setup1 : LastAdaptive",
                             [
                                 "Context:=", "Infinite Sphere1"
                             ],
                             local_var_list,
                             [
                                 "Theta Component:=", "Theta",
                                 "Phi Component:="	, "Phi",
                                 "Mag Component:="		, ["dB(GainTotal)"]
                             ])
        pass
    else:  # 存在就仅更新报告
        oModule.UpdateReports("Gain 3D Radiation Pattern Plot 1")
        pass
    # print("Gain 3D 辐射图 创建完成")
    project_path = oProject.GetPath()  # 获得当前工程文件路径
    project_name = oProject.GetName()   # 获得当前工程名
    data_path = project_path+"/"+project_name+".system_analysis_data"
    if os.path.exists(data_path) is False:
        os.makedirs(data_path)  # 如果不存在则需创建文件夹
        pass
    # 导出数据文件
    oModule.ExportToFile("Gain 3D Radiation Pattern Plot 1", data_path+"/Gain 3D Radiation Pattern Plot 1.csv", False)
    # print("Gain 3D Radiation Pattern Plot 1  数据导出成功")
    pass


def GetAntennaPerformance(frequency: float):
    """
    获得对应频率下的天线参数
    :frequency: 单位 GHz
    :return: 返回天线的S11,带宽,轴比,增益 __S11, __BW, __AR, __Gain
    单位[dB] [GHz] [dB] [dB]
    """
    # 创建变量
    __S11, __BW, __AR, __Gain = 0.0, 0.0, 0.0, 0.0
    # 更新图表文件
    GenerateS11Graph()
    GenerateARBWGraph()
    # Generate3DGainRadiationPattern(2.5)
    GenerateRadiationPattern(2.5)

    data_sheet_name_list = ["S Parameter Plot 1.csv", "Axial Ratio BW Plot 1.csv",
                            "Gain 2D Radiation Pattern Plot 1.csv", "Gain 3D Radiation Pattern Plot 1.csv"]
    project_path = oProject.GetPath()  # 获得当前工程文件路径
    project_name = oProject.GetName()   # 获得当前工程名
    data_path = project_path+"/"+project_name+".system_analysis_data"

    # 使用 pandas 读取 S Parameter Plot 1.csv 文件
    df = pd.read_csv(data_path + "/" + data_sheet_name_list[0])
    # 查询 frequency频率 在仿真文件中是否存在
    if frequency in df["Freq [GHz]"].values:
        # 默认返回10点值
        __index = df[df["Freq [GHz]"] == 2.5].index[0]
        __data = df["dB(S(1,1)) []"].iloc[max(__index-5, 0): min(__index+5, len(df))].values
        __S11 = sum(__data)
        __S11_fre = df[df["Freq [GHz]"] == frequency].values[0][1]
        __S11 = (__S11_fre + __S11)/(len(__data)+1)
        pass
    else:
        __S11 = 0
        __BW = 0

    # 获取 BW 值
    # S11 带宽阈值设置为 -10dB
    threshold_value = -10
    mask = df['dB(S(1,1)) []'] < threshold_value
    groups = (mask != mask.shift()).cumsum()
    result = df[mask].groupby(groups)['Freq [GHz]'].agg(['count', 'first', 'last']).values
    for i in range(len(result)):
        if result[i][1] <= frequency <= result[i][2]:
            __BW = result[i][2] - result[i][1]
            break

    # 获取frequency处 phi theta角为0deg处的 AR 值
    df = pd.read_csv(data_path + "/" + data_sheet_name_list[1])
    if frequency in df["Freq [GHz]"].values:
        __AR = df[df["Freq [GHz]"] == frequency].values[0][3]
        pass
    else:
        __AR = 999

    # 获取frequency处 天线的最大增益
    df = pd.read_csv(data_path + "/" + data_sheet_name_list[2])
    __Gain = df[(df["Freq [GHz]"] == frequency) & (df["Phi [deg]"] == 0) & (df["Theta [deg]"] == 0)].values[0][3]

    return __S11, __BW, __AR, __Gain

if __name__ == '__main__':
    # from hfss import basic
    # projectPath = "P:\\Ansoft\\ProjectStudy.aedt"
    # basic.OpenProject(projectPath)
    # Analyze("Setup1")

    # oProject.handle = oDesktop.GetActiveProject()
    # oDesign.handle = oProject.GetDesign('HFSSDesign12')
    # data_sheet_name_list = ["S Parameter Plot 1.csv", "Axial Ratio BW Plot 1.csv",
    #                         "Gain 2D Radiation Pattern Plot 1.csv", "Gain 3D Radiation Pattern Plot 1.csv"]
    # project_path = oProject.GetPath()  # 获得当前工程文件路径
    # project_name = oProject.GetName()   # 获得当前工程名
    # data_path = project_path +"/" +project_name +".system_analysis_data"
    #
    # # 使用 pandas 读取 S Parameter Plot 1.csv 文件
    # df = pd.read_csv(data_path + "/" + data_sheet_name_list[0])
    # index = df[df["Freq [GHz]"] == 2.5].index[0]
    # print(index)
    # __S11 = df["dB(S(1,1)) []"].iloc[max(index-5, 0): min(index+5, len(df))].values
    # print(__S11)

    # print(oDesign.GetVariables())
    # project_path = oProject.GetPath()  # 获得当前工程文件路径
    # project_name = oProject.GetName()   # 获得当前工程名
    # data_path = project_path+"/"+project_name+".system_analysis_data"
    #
    # df = pd.read_csv(data_path+"/S Parameter Plot 1.csv")
    # print(type(df[df['Freq [GHz]'] == 2.5]))
    # a = df[df['Freq [GHz]'] == 2.5].values[0][1]
    # print(a)

    # oDesign.GetVariables()
    # print(oProject.GetPath())
    # # Analyze('Setup1')
    # oModule.handle = oDesign.GetModule("ReportSetup")
    #
    # print(oDesign.GetName())
    # print(type(oModule.GetAllReportNames()))
    # GenerateS11Graph()
    # GenerateARBWGraph()
    # Generate3DGainRadiationPattern(2.5)
    # GenerateRadiationPattern(2.5)
    pass
