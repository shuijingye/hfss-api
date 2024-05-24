"""
@FileName: hfss/__init__.py
@Description: 该文件主要包含 hfss 接口的对象类的定义
@Author: SJY
@Time: 2023/10/17 15:46
@Coding: UTF-8
@Vision: V1.0
"""

# Official Standard Library
import time, os

# Third-party Library
import win32com.client

# Import the OSL with "from"
from typing import TypeVar
from typing import List, Tuple, Dict

# Import the TPL with "from"
# from win32com import client

# 定义泛型类型变量
# 该变量类型表示 hfss 控件句柄
handle = TypeVar("handle")

# hfss 类的定义: AnsoftApp, Desktop, Project, Design, Editor, Module.
# Object 类的定义 该类包含 常用对象的公用方法
class Object(object):
    def __init__(self):
        self.handle = None

    def GetName(self):
        """ 描述
        Returns the name of the object.
        """
        return self.handle.GetName()

    def GetChildTypes(self):
        return self.handle.GetChildTypes()

    def GetChildNames(self, childType):
        return self.handle.GetChildNames(childType)

    def GetChildObject(self, obj_path):
        return self.handle.GetChildObject(obj_path)

    def GetArrayVariables(self):
        """
        Returns a list of array variables. To get a list of indexed project variables, execute this command using
        oProject. To get a list of indexed local variables, use oDesign.
        :return: Array of strings containing names of variables.
        """
        return self.handle.GetArrayVariables()

    def GetProperties(self, PropTab, PropServer):
        """
        Gets a list of all the properties belonging to a specific <PropServer> and <PropTab>. This can be executed by
        the oProject, oDesign, or oEditor variables.
        :param PropTab: str
            One of the following, where tab titles are shown in parentheses:
            PassedParameterTab ("Parameter Values")
            DefinitionParameterTab (Parameter Defaults")
            LocalVariableTab ("Variables" or "Local Variables")
            ProjectVariableTab ("Project variables")
            ConstantsTab ("Constants")
            BaseElementTab ("Symbol" or "Footprint")
            ComponentTab ("General")
            Component("Component")
            CustomTab ("Intrinsic Variables")
            Quantities ("Quantities")
            Signals ("Signals")
        :param PropServer: str
            An object identifier, generally returned from another script method, such as CompInst@R;2;3
        :return: Array of strings containing the names of the appropriate properties.
        """
        return self.handle.GetProperties(PropTab, PropServer)

    def GetPropertyValue(self, PropTab, PropServer, PropName):
        """
        Returns the value of a single property belonging to a specific <PropServer> and <PropTab>. This function is
        available with the Project, Design or Editor objects, including definition editors.
        :param PropTab: str
            One of the following, where tab titles are shown in parentheses:
                PassedParameterTab ("Parameter Values")
                DefinitionParameterTab (Parameter Defaults")
                LocalVariableTab ("Variables" or "Local Variables")
                ProjectVariableTab ("Project variables")
                ConstantsTab ("Constants")
                BaseElementTab ("Symbol" or "Footprint")
                ComponentTab ("General")
                Component("Component")
                CustomTab ("Intrinsic Variables")
                Quantities ("Quantities")
                Signals ("Signals")
        :param PropServer: str
            An object identifier, generally returned from another script method, such as CompInst@R;2;3
        :param PropName: str
            Name of the property.
        :return: String value of the property.
        """
        return self.handle.GetPropertyValue(PropTab, PropServer, PropName)

    def GetVariables(self):
        """
        Returns a list of all defined variables. To get a list of project variables, execute this command using
        oProject. To get a list of local variables, use oDesign.
        :return: Array of strings containing the variables
        """
        return self.handle.GetVariables()

    def GetVariablesValue(self, VarName):
        """
        Gets the value of a single specified variable. To get the value of project variables, execute this command using
        oProject. To get the value of local variables, use oDesign.
        :param VarName: str
            Name of the variable to access.
        :return: A string representing the value of the variable.
        """
        return self.handle.GetVariablesValue(VarName)

    def ChangeProperty(self, args):
        """

        :param args:
        [ "NAME:AllTabs",
            [ "NAME:LocalVariableTab", ["NAME:PropServers", "LocalVariables"],
            [ "NAME:NewProps",  ["NAME:Width_Ground", "PropType:=" , "VariableProp", "UserDef:=" , True,
                                 "Value:=", "41.72mm" ]]
        ]

        :return: None
        """
        self.handle.ChangeProperty(args)

    def SetPropertyValue(self, propTab, propServer, propName, propValue):
        """
        Sets the value of a single property belonging to a specific PropServer and PropTab. This function is available
        with the Project, Design or Editor objects, including definition editors.This is not supported for properties of
        the following types: ButtonProp, PointProp, V3DPointProp, and VPointProp. Only the ChangeProperty command can be
        used to modify these properties.
        :param propTab: str
            One of the following, where tab titles are shown in parentheses:
                PassedParameterTab ("Parameter Values")
                DefinitionParameterTab (Parameter Defaults")
                LocalVariableTab ("Variables" or "Local Variables")
                ProjectVariableTab ("Project variables")
                ConstantsTab ("Constants")
                BaseElementTab ("Symbol" or "Footprint")
                ComponentTab ("General")
                Component("Component")
                CustomTab ("Intrinsic Variables")
                Quantities ("Quantities")
                Signals ("Signals")
        :param propServer: str
            An object identifier, generally returned from another script method, such as CompInst@R;2;3
        :param propName:
            Name of the property.
        :param propValue:
            The value for the property
        :return: None
        """
        self.handle.SetPropertyValue(propTab, propServer, propName, propValue)

    def SetVariableValue(self, VarName, VarValue):
        """
        Sets the value of a variable. To set the value of a project variable, execute this command using oProject.
        To set the value of a local variable, use oDesign.
        :param VarName: str
            Variable name.
        :param VarValue: value
            New value for the variable.
        :return: None
        """
        self.handle.SetVariableValue(VarName, VarValue)


# Desktop 类的定义 该类对象用于执行桌面级的操作，包括项目管理
class Desktop(object):
    def __init__(self, ansoft_handle):
        self.handle = ansoft_handle.GetAppDesktop()

    def AddMessage(self, project_name, design_name, severity, msg,
                   *category):
        """ 描述
        Add a message with severity and context to message window.

        Parameters
        ----------
        project_name : str
            Project name, an empty string adds the message as the desktop global message
        design_name : str
            Design name, is ignored if project name is empty; an empty string adds the message to
            project node in the message tree
        severity : int
            One of "Error", "Warning" or "Info". Anything other than the first two is treated as "Info"
            0 = Informational, 1 = Warning, 2 = Error, 3 = Fatal
        msg : str
            The message for the message window
        category: str
            Optional. The category is created with the message under the design tree node if the category does not
            exist. If the category already exists, the new message is added to the end of the existing category. It
            is ignored if the project or design is empty. If missing or empty, the message is added to the Design
            node in the message tree.
        Returns
        -------
        None
        """
        self.handle.AddMessage(project_name, design_name, severity, msg, category)

    def CloseProject(self, project_name):
        self.handle.CloseProject(project_name)

    def DeleteProject(self, project_name):
        self.handle.DeleteProject(project_name)

    def EnableAutoSave(self, enable):
        """ 描述
        ----------
        Enable or disable auto-save feature.

        Parameters
        ----------
        enable : bool
           True enables auto-save and False disables it.

        Returns
        -------
        None
        """
        self.handle.EnableAutoSave(enable)

    def GetProjects(self) -> List[handle]:
        """
            Return a list of all the projects that are currently open in the Desktop. Once you have the projects
        you can iterate through them using standard VBScript methods. See the example query.
        :return:
        """
        return self.handle.GetProjects

    def GetProjectList(self) -> List[str]:
        """
        Returns a list of all projects that are open in the Desktop.
        :return:
        """
        return self.handle.GetProjectList()

    def GetActiveProject(self) -> handle:
        """ 描述
        ----------
        Obtain the object for the project that is currently active in the Desktop.
        Note:
        GetActiveProject returns normally if there are no active objects.

        Parameters None
        ----------

        Returns
        -------
        Object, the project that is currently active in the desktop
        """
        __handle = self.handle.GetActiveProject()
        if __handle is None:
            print("当前桌面无可用项目")
        return __handle

    def GetMassages(self, project_name, design_name, severity):
        """ 描述
        Get the messages from a specified project and design

        Parameters
        ----------
        project_name : str
            Name of the project for which to collect messages. An incorrect project name results in no messages
            (design is ignored). An empty project name results in all messages (design is ignored)
        design_name : str
            Name of the design in the named project for which to collect messages. An incorrect design name results
            in no messages for the named project. An empty design name results in all messages for the named project
        severity : int
            Severity is 0-3, and is tied in to info/warning/error/fatal types as follows:
            1 is warning and above
            2 is error and fatal
            3 is fatal only (rarely used)
        Returns
        -------
        A simple array of strings.
        """
        return self.handle.GetMassages(project_name, design_name, severity)

    def NewProject(self):
        self.handle.NewProject()

    def OpenProject(self, file_name: str) -> handle:
        """ 描述
        Get the messages from a specified project and design

        Parameters
        ----------
        file_name : str
        Full path of the project to open

        Returns
        -------
        An object reference to the newly opened project.
        """
        return self.handle.OpenProject(file_name)

    def PauseScript(self, message):
        """ 描述
        Pause the execution of the script and pop up a message to the user. The script execution will not resume
        until the user chooses

        Parameters
        ----------
        message : str
        Any Text

        Returns
        -------
        None
        """
        self.handle.PauseScript(message)

    def QuitApplication(self):
        self.handle.QuitApplication()

    def SetActiveProject(self, project_name):
        """ 描述
        Specify the name of the project that should become active in the desktop. Return that project

        Parameters
        ----------
        project_name : str
        The name of the project already in the Desktop that is to be activated.

        Returns
        -------
        Object, the project that is activated
        """
        return self.handle.SetActiveProject(project_name)

    def SetProjectDirectory(self, directory_path):
        """ 描述

        Parameters
        ----------
        directory_path : str
        The path to the Project directory. This should be writeable by the user

        Returns
        -------
        None
        """
        self.handle.SetProjectDirectory(directory_path)


# Project 类的定义 该类对象对应于在电子桌面中打开的一个项目。它被用于操作项目及其数据。它的数据包括变量、材料定义和一个或多个设计。
class Project(Object):
    def __init__(self):
        # 调用Project 父类的初始化方法 在 python3.x 等同 super.__init__()
        super(Project, self).__init__()
        self.handle = None

    def AnalyzeAll(self):
        """ 描述
        Runs the project-level script command from the script, which simulates all solution setups and Optimetrics
        setups for all design instances in the project. The UI waits until simulation is finished before continuing
        with the script.

        Parameters None
        ----------

        Returns
        -------
        None
        """
        self.handle.AnalyzeAll()

    def CopyDesign(self, design_name):
        return self.handle.CopyDesign(design_name)

    def Paste(self):
        """
        Pastes a design that has already been copied to the clipboard into another design.
        :param pasteOption: {int}
            One of the following:
            0 – Link to the existing design that was copied
            1 – Create a new copy of the design and keep the original layers of the design being copied
            2 – Create a new copy of the design and merge the layers of the design being copied.
        :return: None
        """
        self.handle.Paste()

    def DeleteDesign(self, design_name):
        """ 描述
        Deletes a specified design in the project.

        Parameters
        ----------
        design_name : str
        The name of the design

        Returns
        -------
        None
        """
        self.handle.DeleteDesign(design_name)

    def GetActiveDesign(self) -> handle:
        """ 描述
        Returns the design in the active project
        Note :
        GetActiveDesign will return normally if there are no active objects.

        Parameters None
        ----------

        Returns
        -------
        The active design
        """
        return self.handle.GetActiveDesign()

    def GetDefinitionManager(self):
        """ 描述
        Gets the DefinitionManager object.

        Parameters None
        ----------

        Returns str
        -------
        Object for the DefinitionManager.
        """
        return self.handle.GetDefinitionManager()

    def GetDesign(self, design_name: str) -> handle:
        """ 描述
        Returns the interface to a specific design in a given project

        Parameters
        ----------
        design_name : str
            Name of the design

        Returns
        -------
        Returns the interface to the specified design.
        """
        return self.handle.GetDesign(design_name)

    def GetPath(self):
        """ 描述
        Returns the location of the project on disk.

        Parameters None
        ----------

        Returns str
        -------
        String containing the path to the project, not including the project name.
        """
        return self.handle.GetPath()

    def GetPropValue(self, prop_path):
        """ 描述
        Returns the property value for the active project object, or specified property values.

        Parameters
        ----------
        prop_path : str
            A child object's property path. See: Property Function Summary.

        Returns
        -------
             The property value of a child object.

        Example
        -------
        oProject.GetPropValue("TeeModel/offset") //get the offset variable value in the TeeModel Design
        oProject.GetPropValue("TeeModel/Results/S ParameterPlot 1/Display Type") // Get the report display type.
        """
        return self.handle.GetPropValue(prop_path)

    def GetTopDesignList(self):
        """ 描述
        Returns a list of top-level design names.

        Parameters None
        ----------

        Returns array
        -------
        An array containing string names of top-level designs.
        """
        return self.handle.GetTopDesignList()

    def InsertDesign(self, design_type, design_name, solution_type):
        """ 描述
        Returns a list of top-level design names.
        Parameters
        ----------
        design_type : str
            Design type. Use "HFSS".
        design_name : str
            Design name.
        solution_type : str
            Can be "DrivenModal", "DrivenTerminal", or "Eigenmode".
        Returns
        -------
        Design handle
        """
        return self.handle.InsertDesign(design_type, design_name, solution_type, "")

    def Redo(self):
        self.handle.Redo()

    def Rename(self, new_name, over_write_ok):
        """ 描述
        Renames the project and saves it. Similar to SaveAs().

        Parameters
        ----------
        new_name : str
            The desired name of the project. The path is optional.
        over_write_ok : bool
            True to overwrite the same named file on disk if it exists. False to prevent overwrite.

        Returns
        -------
        None
        """
        self.handle.Rename(new_name, over_write_ok)

    def Save(self):
        self.handle.Save()

    def SaveAs(self, new_name, over_write_ok, default_action="", *overwrite_actions):
        """ 描述
        Saves the project under a new name. Requires a full path. Note that this script takes two parameters for
        non-schematic/layout designs and four parameters for schematic/layout designs.

        Parameters
        ----------
        new_name : str
            The desired name of the project, with directory and extension.
        over_write_ok : bool
            True to overwrite the file of the same name, if it exists. False to prevent overwrite.
        default_action : str
            For Schematic/Layout projects only. Otherwise omit. See note below.
            Valid actions: ef_overwrite , ef_copy_no_overwrite, ef_make_path_absolute, or empty string.
        overwrite_actions : Array
            For Schematic/Layout projects only. Otherwise omit. See note below.
            Structured array:Array("Name: <Action>",<FileName>, <FileName>, ...)
            Valid actions: ef_overwrite , ef_copy_no_overwrite, ef_make_path_absolute, or empty string.

        Returns
        -------
        None
        """
        if default_action == "" and overwrite_actions is None:
            pass
        self.handle.SaveAs(new_name, over_write_ok, "", "")

    def SetActiveDesign(self, design_name):
        """ 描述
        Sets a design to be the active design.

        Parameters
        ----------
        design_name : str
            Name of the design to set as the active design.

        Returns
        -------
        None
        """
        self.handle.SetActiveDesign(design_name)

    def SetPropValue(self, prop_path, value):
        """ 描述
        Sets the property value for the active property object.

        Parameters
        ----------
        prop_path : str
            A child object's property path. See: PropertyFunction.

        value : str
            New property value

        Returns
        -------
        bool
        True if the property is found and the new value is valid. Otherwise, False.
        """
        return self.handle.SetPropValue(prop_path, value)

    def SimulateAll(self):
        """ 描述
        Simulates all solution setups and Optimetrics setups for all design instances in the project. Script
        processing only continues when all analyses are finished.

        Parameters None
        ----------

        Returns
        -------
        None
        """
        self.handle.SimulateAll()

    def Undo(self):
        """ 描述
        Cancels the last project-level command.

        Parameters None
        ----------

        Returns
        -------
        None
        """
        self.handle.Undo()

    def UpdateDefinitions(self):
        """ 描述
        Updates all definitions. The Messages window reports when definitions are updated, or warns
        when definitions cannot be found.

        Parameters None
        ----------

        Returns
        -------
        None
        """
        self.handle.UpdateDefinitions()

    def ValidateDesign(self):
        """ 描述
        Returns whether a design is valid.

        Parameters None
        ----------

        Returns int
        -------
        1 – Validation passed.
        0 – Validation failed.
        """
        return self.handle.ValidateDesign()

    #  Dataset Script Commands
    def AddDataSet(self, dataset_data_array):
        """ 描述
        Adds a dataset.

        Parameters
        ----------
        dataset_data_array : Array
            ["NAME:<DatasetName>",["NAME:Coordinates",<CoordinateArray>,<CoordinateArray>, ...]]
        DatasetName str
            Name of the dataset
        CoordinateArray Array
            ["NAME:Coordinate", "X:=", <double>,"Y:=",<double>]

        Returns
        -------
            None
        """
        self.handle.AddDataset(dataset_data_array)

    def DeleteDataset(self, dataset_name):
        """ 描述
        Deletes a specified dataset.

        Parameters
        ----------
        dataset_name : str
            Name of the dataset found in the project.

        Returns
        -------
            None
        """
        self.handle.DeleteDataset(dataset_name)

    def EditDataset(self, original_name, dataset_data_array):
        """ 描述
        Modifies a dataset. When a dataset is modified, its name as well as its data can be changed.

        Parameters
        ----------
        original_name : str
            Name of the dataset before editing.
        dataset_data_array : Array
            ["NAME:<DatasetName>",["NAME:Coordinates",<CoordinateArray>,<CoordinateArray>, ...]]
        DatasetName str
            Name of the dataset
        CoordinateArray Array
            ["NAME:Coordinate", "X:=", <double>,"Y:=",<double>]

        Returns
        -------
            None
        """
        self.handle.EditDataset(original_name, dataset_data_array)

    def ImportDataset(self, dataset_file_full_path):
        """ 描述
        Imports a dataset from a named file.

        Parameters
        ----------
        dataset_file_full_path : str
            The full path to the file containing the dataset values.

        Returns
        -------
            None

        Example
        oProject.ImportDataset('e:\\mp\\dataset.txt')
        """
        self.handle.ImportDataset(dataset_file_full_path)


# Design 类的定义 该类对象与项目中的一个设计相对应。此对象用于操作设计及其数据，包括变量、模块和编辑器。
class Design(Object):
    def __init__(self):
        super(Design, self).__init__()
        self.handle = None

    """ < ModuleName >
    Analysis Module – "AnalysisSetup"
    Boundary Module – "BoundarySetup"
    Field Overlays Module – "FieldsReporter"
    Mesh Module – "MeshSetup"
    Optimetrics Module – "Optimetrics"
    Radiation Module – "RadField"
    Radiation Setup Manager Module – "RadiationSetupMgr"
    Reduce Matrix Module – "ReduceMatrix"
    Reporter Module – "ReportSetup"
    Simulation Setup Module – "SimSetup"
    Solutions Module – "Solutions"
    """

    def AddDesignVariablesForDynamicLink(self, InstanceID):
        """
        Creates design variables based on any in a Circuit, HFSS, 3D Layout, Q3D, or 2D extractor project that is
        connected to a Circuit or 3D Layout design through a dynamic link.
        :param InstanceID: str
            Instance ID for the dynamic link component
        :return: None
        """
        self.handle.AddDesignVariablesForDynamicLink(InstanceID)

    def ApplyMeshOps(self, setup_name_array):
        """ 描述
        If any mesh operations were defined and not yet performed in the current variation for the specified solution
        setups, they will be applied to the current mesh. If necessary, an initial mesh will be computed first.
        No further analysis will be performed.

        Parameters
        ----------
        setup_name_array : Array
            Array containing the string names of setups.

        :return: int
            0 – Success
            Else – Error
        Example
        oDesign.ApplyMeshOps(['Setup1','Setup2'])
        """
        return self.handle.ApplyMeshOps(setup_name_array)

    def Analyze(self, setup):
        """ 描述
        Solves a single solution setup and all of its frequency sweeps.

        Parameters
        ----------
        setup : str
            Setup name

        Returns
        -------
            int
            0 – Success
            Else – Error
        Example
        oDesign.Analyze('TR1')
        """
        return self.handle.Analyze(setup)

    def AnalyzeDistributed(self, SetupName):
        """
        Performs a distributed analysis.
        :param SetupName: str
            String name of the setup to be analyzed.
        :return: int
            0 – Success, Else – Error
        """
        self.handle.AnalyzeDistributed(SetupName)

    def ClearLinkedData(self):
        """
        Clear the linked data of all the solution setups. Similar to the ClearLinkedData command for the module level.
        :return: None
        """
        self.handle.ClearLinkedData()

    def ConstructVariationString(self, ArrayOfVariableNames, ArrayOfVariableValuesIncludingUnits):
        """
        Lists and orders the variables and values associated with a design variation.
        :param ArrayOfVariableNames: Array of Strings
            List of variable names.
        :param ArrayOfVariableValuesIncludingUnits: Array of Strings
            List of variable values,including units, in the same order as the list of names
        :return: Returns variation string with the variables ordered to correspond to the order of variables in design
        variations. The values for the variables are inserted into the variation string. For an example of how
        ConstructionVariationString can be used, see the Python example.

        varStr = oDesign.ConstructVariationString(["xx", "yy"],["2mm", "1mm"])
        oDesign.ExportProfile("Setup1", varStr, "C:\\profile.prof")
        """
        return self.handle.ConstructVariationString(ArrayOfVariableNames, ArrayOfVariableValuesIncludingUnits)

    def DeleteFieldVariation(self, variationKeys, deleteMesh, deleteLinkedData, preserveRadFields):
        """
        Deletes field variations, fields and meshes, or fields and meshes for specified variations.
        :param variationKeys: Array
            Array containing either "All" string to select all variations, or strings of variations to include.
        :param deleteMesh: Boolean
            When true, deletes mesh data.
        :param deleteLinkedData: Boolean
            When true, deletes linked data.
        :param preserveRadFields: Boolean
            Optional. When true, preserves radiated fields data.
        :return: None
        """
        self.handle.DeleteFieldVariation(variationKeys, deleteMesh, deleteLinkedData, preserveRadFields)

    def DeleteFullVariation(self, variationKeys, deleteLinkedData):
        """
        Deletes either all solution data, or selected variation data.
        :param variationKeys: array
            Array containing either "All" string to select all variations, or strings of variations to include.
        :param deleteLinkedData: bool
            When true, deletes linked data.
        :return: None
        """
        self.handle.DeleteFullVariation(variationKeys, deleteLinkedData)

    def DeleteLinkedDataVariation(self, DesignVariationKeys):
        """
        Deletes the linked data of specified variations.
        :param DesignVariationKeys: Array
            Array containing strings of the variations keys whose linked data are going to be deleted.
        :return: None
        """
        self.handle.DeleteLinkedDataVariation(DesignVariationKeys)

    def EditDesignSettings(self, overrideArray):
        """
        Edits design settings.
        :param overrideArray: Structured array.
            For HFSS design:
            Array( "NAME:Design Settings Data",
                   "Use Advanced DC Extrapolation:=", <Boolean>,
                   "Port Validation Settings=:", ["Standard" | "Extended"]
                   "Calculate Lossy Dielectrics:=", <Boolean> )
            "Use Power S:=", <Boolean>,
            "Export After Simulation:=", <boolean>,
            "Export Dir:=", "<path>",
            "Allow Material Override:=", <boolean>,
            "Calculate Lossy Dielectrics:=", <boolean>,
            "EnabledObjects:=", Array(),
            "Model validation Settings:=",
                Array("NAME:Validation Options",
                      "EntityCheckLevel:=", "[strict | none | warningOnly | basic ]",
                      "IgnoreUnclassifiedObjects:=", <boolean>,
                      "SkipIntersectionChecks:=", <boolean>),
                      "Design Validation Settings:=", "Perform full validations" )
        :return: None
        """
        self.handle.EditDesignSettings(overrideArray)

    def EditNotes(self, DesignNotes):
        """
        Update the notes for the design.
        :param DesignNotes: str
            New design notes that are applied, replacing any current notes.
        :return: None
        """
        self.handle.EditNotes(DesignNotes)

    def ExportConvergence(self, Setup, VariationKeys, FilePath, OverwriteIfExists):
        """
        For a given variation, exports convergence data (max mag delta S, E, freq) to *.conv file.
        :param Setup: str
            Setup name.
        :param VariationKeys: str
            The variation. Pass empty string for the current nominal variation.
        :param FilePath: str
            Full path to desired *.conv file location.
        :param OverwriteIfExists: Boolean
            Optional. If True, overwrites any existing file at the same path.
        :return: None
        """
        self.handle.ExportConvergence(Setup, VariationKeys, FilePath, OverwriteIfExists)

    def ExportMatrixData(self, ParametersArray):
        """

        :param ParametersArray:
        [
            <FileName>: <str> The path and name of the file where the data will be exported. The file extension
            determines the file format.
            <SolnType>: <str> One of the following "C", "AC RL" and "DC RL"
            <DesignVariationKey>: <str> Design variation string
            Example:"radius = 3mm" The empty variation string ("") is interpreted to mean the current nominal variation.
            <Solution> <SolveSetup>:<Soln>
                <SolveSetup>: <str> Name of the solve setup
                <Soln>: <str> Name of the solution at a certain adaptive pass.
            <Matrix>: <str> Either "Original" or one of the reduce matrix setup names
            <ResUnit>: <str> Unit used for the resistance value
            <IndUnit>: <str> Unit used for the inductance value
            <CapUnit>: <str> Unit used for the capacitance value
            <CondUnit>: <str> Unit used for the conductance value
            <Frequency>: <double> Frequency in hertz.
            <MatrixType>: <str> Value: "Maxwell", "Spice", "Couple" , (one or all of these). Matrix type to export.
            <PassNumber>: <integer> Pass number.
            <ACPlusDCResistance>: <bool> Default Value: false
        ]
        :return: None
        """
        self.handle.ExportMatrixData(ParametersArray)

    def ExportMeshStats(self, setupName, variationString, filePath, overwriteIfExists):
        """
        Exports mesh statistics to a file.
        :param setupName: str
            Name of the solution setup.
        :param variationString: str
            Name of the variation. An empty string is interpreted as the current nominal variation.
        :param filePath: str
            Full file path, including extension *.ms.
        :param overwriteIfExists: bool
            Optional. Defaults to True. If True, script overwrites any existing file of the same name.
            If False, no overwrite.
        :return: None
        """
        self.handle.ExportMeshStats(setupName, variationString, filePath, overwriteIfExists)

    def ExportNMFData(self, SolutionName, FileName, ReduceMatrix, ReferenceImpedance, FrequencyArray, DesignVariation,
                      Format, Length, PassNumber):
        """
        Exports s-parameters in neutral file format.
        :param SolutionName: <str>
            Format: <SetupName>:<SolutionName>
            Solution that is exported.
        :param FileName: <str>
            The name of the file. The file extension will determine the file format.
        :param ReduceMatrix: <str>
            Either "Original" or one of the reduce matrix setup name.
        :param ReferenceImpedance: <Double>
            Reference impedance
        :param FrequencyArray: <Array of doubles>
            Frequency points in the sweep that is exported.
        :param DesignVariation: <str>
            Design variation at which the solution is exported.
        :param Format: <str> : "MagPhase", "RealImag", "DbPhase"
            The format in which the s-parameters will be exported.
        :param Length: <str>
            Length for exporting s-parameters.
        :param PassNumber: <int>
            Pass number.
        :return: None
        """
        self.handle.ExportNMFData(SolutionName, FileName, ReduceMatrix, ReferenceImpedance, FrequencyArray,
                                  DesignVariation, Format, Length, PassNumber)

    def ExportProfile(self, SetupName, VariationString, filePath, overwriteIfExists=True):
        """
        Exports a solution profile to file.
        :param SetupName: {str}
            Text of the design's notes.
        :param VariationString: (str)
            Variation name. Pass empty string for the current nominal variation.
        :param filePath: (str)
            Full file path, including extension *.prof.
        :param overwriteIfExists: (bool)
            Optional. Default is True. If True, overwrites any existing file of the same name. If False, no overwrite.
        :return: None
        """
        self.handle.ExportProfile(SetupName, VariationString, filePath, overwriteIfExists)

    def GenerateMesh(self, SimulationNames):
        """
        If any mesh operations have been defined but not yet performed in the current variation for the specified
        solution setups, they will be applied to the current mesh. If necessary, an initial mesh will be computed first.
        No further analysis will be performed.
        :param SimulationNames: {Array}
            List of simulation names
        :return: {int} 0 – Success, Else – Failure
        """
        self.handle.GenerateMesh(SimulationNames)

    def GetDesignType(self):
        """
        Returns the design type of the active design.
        :return: String indicating the design type of the active design ("Circuit Design", "Circuit Netlist","EMIT",
        "HFSS 3D Layout Design", "HFSS", "HFSS-IE", "Icepak", "Maxwell 2D", "Maxwell 3D", "Q2D Extractor",
         "Q3DExtractor", "RMxprt", or"Twin Builder").
        """
        return self.handle.GetDesignType()

    def GetDesignValidationInfo(self, validationResultFilePath):
        """
        :param validationResultFilePath: {str}
            Full path for validation file, including *.xml extension.
        :return:
            0 – Failure(validation not done); 1 – Success (results written to specified location; see file format below)
            <ValidationInfo>
            <DesignValid>false</DesignValid>
            <IgnoreUnclassifiedObjects>false</IgnoreUnclassifiedObjects>
            <SkipIntersectionChecks>false</SkipIntersectionChecks>
            <EntityCheckLevel>Strict</EntityCheckLevel>
            <ValidationLevel>Perform full validations with standard port
            validations</ValidationLevel>
            <Messages>
            <Message>[warning] Warning - Boundary &apos;Rad1&apos; and
            Boundary &apos;PerfE1&apos; overlap. </Message>
            <Message>[error] Error - Port &apos;Port1&apos; and Port
            &apos;1&apos; overlap. </Message>
            </Messages>
            </ValidationInfo>
        """
        return self.handle.GetDesignValidationInfo(validationResultFilePath)

    def GetEditSourcesCount(self):
        """
        For Characteristic Mode Analysis(CMA), returns the number of sources listed in the Edit Sources panel.
        :return: INT number of sources.
        """
        return self.handle.GetEditSourcesCount()

    def GetManagedFilesPath(self):
        """
        Get the path to the project's results folder
        :return: String containing path where project results are located,
        such as C:/Users/[username]/Ansoft/Project31.aedtresults/mf_0/
        """
        return self.handle.GetManagedFilesPath()

    def GetModule(self, moduleName: str) -> handle:
        """
        Returns the IDispatch for the specified module.
        :param moduleName: str
        :return: Module IDispatch
        """
        return self.handle.GetModule(moduleName)

    def GetNominalVariation(self):
        """
        Returns the current nominal variation.
        :return: String containing current nominal variation.
        """
        return self.handle.GetNominalVariation()

    def GetNoteText(self):
        """
        Return the text of the note attached to a design.
        :return: Returns text of the design note.
        """
        return self.handle.GetNoteText()

    def GetPostProcessingVariables(self):
        """
        Returns the list of post-processing variables.
        :return: Returns array containing variables.
        """
        return self.handle.GetPostProcessingVariables()

    def GetPropValue(self, propPath):
        """
        Returns the property value for the active design object, or specified property values.
        :param propPath: A child object's property path. See:Object Property Script Function Summary.
        :return: The property value (integer, string, or array of strings) of the specified child object.
        """
        return self.handle.GetPropValue(propPath)

    def GetSolutionType(self):
        """
        Returns the design's solution type.
        :return: String containing the solution type.
        Examples of possible values are: "DrivenModal", "DrivenTerminal", "Eigenmode",
        "Transient" or "Transient Network".
        """
        return self.handle.GetSolutionType()

    def GetVariationVariableValue(self, VariationString, VariableName):
        """
        Returns the value for a specified variation's variable.
        :param VariationString:
            The name of the design variation.
        :param VariableName:
            The name of the variable.
        :return: Returns a double precision value in SI units, interpreted to mean the value of the variable contained
        in the variation string.
        """
        return self.handle.GetVariationVariableValue(VariationString, VariableName)

    def PasteDesign(self, pasteOption):
        """
        Pastes a design that has already been copied to the clipboard into another design.
        :param pasteOption: {int}
            One of the following:
            0 – Link to the existing design that was copied
            1 – Create a new copy of the design and keep the original layers of the design being copied
            2 – Create a new copy of the design and merge the layers of the design being copied.
        :return: None
        """
        self.handle.PasteDesign(pasteOption)

    def Redo(self):
        self.handle.Redo()

    def RenameDesignInstance(self, OldName, NewName):
        self.handle.RenameDesignInstance(OldName, NewName)

    def SetActiveEditor(self, EditorName):
        """
        Sets the active editor.
        :param EditorName: Text of the design's notes.
        :return: Editor object.
        """
        return self.handle.SetActiveEditor(EditorName)

    def SetBackgroundMaterial(self, matName):
        """
        Sets the design's background material.
        :param matName: Material name.
        :return: None
        """
        self.handle.SetBackgroundMaterial(matName)

    def SetDesignSettings(self, OverrideArray):
        """
        To set the design settings.
        :param OverrideArray:
        Array("NAME:Design Settings Data",
        "Use Advanced DC Extrapolation:= <Boolean>,
        "Port Validation Settings=: ["Standard" | "Extended"]
        "Calculate Lossy Dielectrics:=", <Boolean>)
        "Use Power S:=", <Boolean>,
        "Export After Simulation:=", <boolean>, "Export Dir:=", "<path>",
        "Allow Material Override:=", <boolean>,
        "Calculate Lossy Dielectrics:=", <boolean>,
        "EnabledObjects:=", Array(),
        "Maximum Number of Frequencies for Broadband Adapt:=", <int>
        ],
        [
        "Model validation Settings:=", Array("NAME:Validation Options",
        "EntityCheckLevel:=", "[strict | none | warningOnly | basic ]",
        "IgnoreUnclassifiedObjects:=", <boolean>,
        "SkipIntersectionChecks:=", <boolean>),
        "Design Validation Settings:=", "Perform full validations" )
        :return: None
        """
        self.handle.SetDesignSettings(OverrideArray)

    def SetLengthSettings(self, DistributedUnits, LumpedLength, RiseTime):
        """
        Sets the distributed and lumped lengths of the design, as well as the rise time.
        :param DistributedUnits: Length units used for post-processing.
        :param LumpedLength: Length of design in post-processing.
        :param RiseTime: Time used in post-processing.
        :return: None
        """
        self.handle.SetLengthSettings(DistributedUnits, LumpedLength, RiseTime)

    def SetObjectTemperature(self, TemperatureSettings):
        """

        :param TemperatureSettings:
            <IncludeTemperatureDependence> : Boolean
                True to include temperature dependence. False to not include.
            <EnableFeedback>
                True to enable feedback. False to not enable feedback.
            <Temperatures>
                Objects and temperatures.
        :return: None
        """
        self.handle.SetObjectTemperature(TemperatureSettings)

    def SetPropValue(self, path, data):
        """
        Sets the property value for the active property object
        :param path: A child object's property path
        :param data: New property value
        :return: True if the property is found and the new value is valid. Otherwise returns False.
        """
        return self.handle.SetPropValue(path, data)

    def Solve(self, SimulationNames):
        """
        Performs one or more simulation. The next script command will not be executed until the simulation(s) are
        complete.
        :param SimulationNames: Array
            Array containing string simulation names.
        :return:  0 – Simulation(s) completed. 1 – Simulation error. -1 – Command execution error.
        """
        return self.handle.Solve(SimulationNames)

    def Undo(self):
        self.handle.Undo()

    def ValidateDesign(self):
        return self.handle.ValidateDesign()

    def ImportIntoReport(self, ReportName, FileName):
        self.handle.ImportIntoReport(ReportName, FileName)


# Module 类的定义 该类对象对应于设计中的一个模块。模块用于处理一组相关的功能。
class Module(Object):
    def __init__(self):
        super(Module, self).__init__()
        self.handle = None

    def AssignDCThickness(self, object_name_array, thickness_value_array, enable_auto):
        """ 描述
        Assigns DC thickness to more accurately compute DC resistance of a thin conducting object for which
        Solve Inside is not selected.

        Parameters
        ----------
        object_name_array : array
            Array containing string object name(s). These objects will have their thickness changed.
        thickness_value_array : array
            Array containing string thickness values, including units. The values are applied to the objects in the
            object_name_array, in order (the first object receives the first value, the second object the second value,
            and so on).
            An empty string leaves the value as-is.
            Alternately, can be string "Infinite" for infinite thickness or "Effective" to have it calculated
            automatically.
        enable_auto : str
            "EnableAuto" or "DisableAuto". Omitting this string leaves the selection unchanged from its current setting.

        Returns
        -------
            None

        Example
        oModule.AssignDCThickness(["pad1", "pin1", "pad2", "pad3", "pad4", "pin2", "pin3", "pin4", "gnd"],
                                  [ "1mm", "\"<Infinite>\"", "\"<Effective>\"",
                                    "2mm", "\"<Infinite>\"", "\"<Effective>\"",
                                    "3mm", "\"<Infinite>\"", "\"<Effective>\"" ],
                                  "EnableAuto" )
        """
        self.handle.AssignDCThickness(object_name_array, thickness_value_array, enable_auto)

    def AssignArray(self, Parameters):
        self.handle.AssignArray(Parameters)

    def DeleteArray(self):
        self.handle.DeleteArray()

    def EditArray(self, Parameters):
        self.handle.EditArray(Parameters)

    def SetSourceContexts(self, sourceID):
        self.handle.SetSourceContexts(sourceID)

    # OutputVariable
    def CreateOutputVariable(self, OutputVarName, Expression, SolutionName, Type, parameterArray):
        """
        Adds a new output variable. Different forms of this command are executed for different design types.

        :param OutputVarName: [str]
            Name of the new output variable.
        :param Expression: [value]
            Value to assign to the variable.
        :param SolutionName: [str]
            Name of the solution, as seen in the output variable UI.
        :param Type: [str] <reportType> or <solutionType>
            The name of the report type as seen in the output variable UI;
            For Layout Editor, use the solution type.
        :param parameterArray: [Array] <ContextArray> or <DomainArray>
            Structured array containing context for which the output variable expression is being evaluated.
                Array("Context:=", <string>)
            For Layout Editor, use Domain array:
                Array("Domain:=", <string>)
        :return: None
        """
        self.handle.CreateOutputVariable(OutputVarName, Expression, SolutionName, Type, parameterArray)

    def DeleteOutputVariable(self, OutputVarName):
        self.handle.DeleteOutputVariable(OutputVarName)

    def DoesOutputVariableExist(self, OutputVarName):
        return self.handle.DoesOutputVariableExist(OutputVarName)

    def EditOutputVariable(self, OutputVarName, NewExpression, NewVarName, SolutionName, Type, parameterArray):
        self.handle.EditOutputVariable(OutputVarName, NewExpression, NewVarName, SolutionName, Type, parameterArray)

    def GetOutputVariableValue(self, OutputVarName, IntrinsicVariation, SolutionName, ReportTypeName, ContextArray):
        return self.handle.GetOutputVariableValue(OutputVarName, IntrinsicVariation, SolutionName, ReportTypeName,
                                                  ContextArray)

    # ReportSetup
    def AddCartesianLimitLine(self, ReportName, Def):
        self.handle.AddCartesianLimitLine(ReportName, Def)

    def AddCartesianXMarker(self, ReportName, MarkerName, XValue):
        self.handle.AddCartesianXMarker(ReportName, MarkerName, XValue)

    def AddCartesianYMarker(self, ReportName, MarkerName, YValue):
        self.handle.AddCartesianYMarker(ReportName, MarkerName, YValue)

    def AddDeltaMarker(self, ReportName, MarkerName1, CurveName1, PrimarySweepValue1, MarkerName2, CurveName2,
                       PrimarySweepValue2):
        self.handle.AddDeltaMarker(ReportName, MarkerName1, CurveName1, PrimarySweepValue1, MarkerName2, CurveName2,
                                   PrimarySweepValue2)

    def AddMarker(self, ReportName, MarkerName, CurveName, PrimarySweepValue):
        self.handle.AddMarker(ReportName, MarkerName, CurveName, PrimarySweepValue)

    def AddNote(self, ReportName, NoteDataArray):
        self.handle.AddNote(ReportName, NoteDataArray)

    def AddTraceCharacteristics(self, ReportName, FunctionName, FunctionArgs, RangeArgs):
        self.handle.AddTraceCharacteristics(ReportName, FunctionName, FunctionArgs, RangeArgs)

    def AddTraces(self, ReportName, SolutionName, ContextArray, FamiliesArray, ReportDataArray):
        self.handle.AddTraces(ReportName, SolutionName, ContextArray, FamiliesArray, ReportDataArray)

    def ClearAllMarkers(self, ReportName):
        self.handle.ClearAllMarkers(ReportName)

    def ClearAllTraceCharacteristics(self, plotName):
        self.handle.ClearAllTraceCharacteristics(plotName)

    def CopyTracesData(self, ReportName, TracesArray):
        self.handle.CopyTracesData(ReportName, TracesArray)

    def CopyReportData(self, ReportsArray):
        self.handle.CopyReportData(ReportsArray)

    def CopyReportDefinitions(self, ReportsArray):
        self.handle.CopyReportDefinitions(ReportsArray)

    def CopyTraceDefinitions(self, ReportName, TracesArray):
        self.handle.CopyTraceDefinitions(ReportName, TracesArray)

    def CreateReport(self, ReportName, ReportType, DisplayType, SolutionName, ContextArray, FamiliesArray,
                     ReportDataArray):
        self.handle.CreateReport(ReportName, ReportType, DisplayType, SolutionName, ContextArray, FamiliesArray,
                                 ReportDataArray, [])

    def CreateReportFromFile(self, FilePathName):
        self.handle.CreateReportFromFile(FilePathName)

    def CreateReportFromTemplate(self, TemplatePath):
        self.handle.CreateReportFromTemplate(TemplatePath)

    def CreateReportOfAllQuantities(self, reportNameArg, reportTypeArg, displayTypeArg, solutionNameArg,
                                    simValueCtxtArg, categoryNameArg, pointSetArg, commonComponentsOfTracesArg,
                                    extTraceInfoArg):
        self.handle.CreateReportOfAllQuantities(reportNameArg, reportTypeArg, displayTypeArg, solutionNameArg,
                                                simValueCtxtArg, categoryNameArg, pointSetArg,
                                                commonComponentsOfTracesArg, extTraceInfoArg)

    def DeleteMarker(self, MarkerName):
        self.handle.DeleteMarker(MarkerName)

    def DeleteAllReports(self):
        self.handle.DeleteAllReports()

    def DeleteReports(self, ReportNameArray):
        self.handle.DeleteReports(ReportNameArray)

    def DeleteTraces(self, TraceSelectionArray):
        self.handle.DeleteTraces(TraceSelectionArray)

    def ExportFieldsToFile(self, ParametersArray):
        self.handle.ExportFieldsToFile(ParametersArray)

    def ExportParametersToFile(self, ParametersArray):
        self.handle.ExportParametersToFile(ParametersArray)

    def ExportImageToFile(self, ParametersArray):
        self.handle.ExportImageToFile(ParametersArray)

    def ExportPlotImageToFile(self, ParametersArray):
        self.handle.ExportPlotImageToFile(ParametersArray)

    def ExportReport(self, ReportName, FileName):
        self.handle.ExportReport(ReportName, FileName)

    def ExportToFile(self, ReportName, FileName, OverWrite=False):
        self.handle.ExportToFile(ReportName, FileName, OverWrite)

    def ExportMarkerTable(self, FileFullPath):
        self.handle.ExportMarkerTable(FileFullPath)

    def FFTOnReport(self, PlotName, FFTWindowType, function):
        self.handle.FFTOnReport(PlotName, FFTWindowType, function)

    def GetAllCategories(self, reportTypeArg, displayTypeArg, solutionNameArg, simValueCtxtArg, categoryNameArray):
        return self.handle.GetAllCategories(reportTypeArg, displayTypeArg, solutionNameArg, simValueCtxtArg,
                                            categoryNameArray)

    def GetAllReportNames(self) -> Tuple[str, ...]:
        """
        返回当前 design 下包含的所有 报告名称的列表
        :return: [str, str, ...]
        """
        return self.handle.GetAllReportNames()

    def GetAllQuantities(self, reportTypeArg, displayTypeArg, solutionNameArg, simValueCtxtArg, categoryNameArg,
                         quantityNameArray):
        return self.handle.GetAllQuantities(reportTypeArg, displayTypeArg, solutionNameArg, simValueCtxtArg,
                                            categoryNameArg, quantityNameArray)

    def GetAvailableDisplayTypes(self, reportTypeArg, displayTypeArray):
        return self.handle.GetAvailableDisplayTypes(reportTypeArg, displayTypeArray)

    def GetAvailableReportTypes(self, reportTypeArray):
        return self.handle.GetAvailableReportTypes(reportTypeArray)

    def GetAvailableSolutions(self, reportTypeArg, solutionArray):
        return self.handle.GetAvailableSolutions(reportTypeArg, solutionArray)

    def GetDisplayType(self, ReportName):
        return self.handle.GetDisplayType(ReportName)

    def GetPropNames(self, bIncludeReadOnly):
        return self.handle.GetPropNames(bIncludeReadOnly)

    def GetPropValue(self, propPath):
        return self.handle.GetPropValue(propPath)

    def GetReportTraceNames(self, plot_name):
        return self.handle.GetReportTraceNames(plot_name)

    def GetSolutionContexts(self, reportTypeArg, displayTypeArg, solutionNameArg, contextNameArray):
        return self.handle.GetSolutionContexts(reportTypeArg, displayTypeArg, solutionNameArg, contextNameArray)

    def GetSolutionDataPerVariation(self, reportTypeArg, solutionNameArg, simValueCtxtArg, familiesArg, expressionArg):
        return self.handle.GetSolutionDataPerVariation(reportTypeArg, solutionNameArg, simValueCtxtArg, familiesArg,
                                                       expressionArg)

    def GroupPlotCurvesByGroupingStrategy(self, reportName, curve_grouping_strategy):
        return self.handle.GroupPlotCurvesByGroupingStrategy(reportName, curve_grouping_strategy)

    def MovePlotCurvesToGroup(self, ReportName, curveN_name, stack_name):
        self.handle.MovePlotCurvesToGroup(ReportName, curveN_name, stack_name)

    def PasteReports(self):
        self.handle.PasteReports()

    def PasteTraces(self, plotName):
        self.handle.PasteTraces(plotName)

    def RenameReport(self, OldReportName, NewReportName):
        self.handle.RenameReport(OldReportName, NewReportName)

    def RenameTrace(self, plotName, traceID, newName):
        self.handle.RenameTrace(plotName, traceID, newName)

    def ResetPlotSettings(self, PlotName):
        self.handle.ResetPlotSettings(PlotName)

    def SavePlotSettingsAsDefault(self, PlotName):
        self.handle.SavePlotSettingsAsDefault(PlotName)

    def SetPropValue(self, proPath, newView):
        self.handle.SetPropValue(proPath, newView)

    def UpdateTraces(self, ReportName, SolutionName, ContextArray, DomainType, FamiliesArray, ValueArray,
                     ReportDataArray, ReportQuantityArray, Array):
        self.handle.UpdateTraces(ReportName, SolutionName, ContextArray, DomainType, FamiliesArray, ValueArray,
                                 ReportDataArray, ReportQuantityArray, Array)

    def UpdateTracesContextandSweeps(self, ReportName, traceIDsArray, SolutionName, ContextArray, pointSetArray):
        self.UpdateTracesContextandSweeps(ReportName, traceIDsArray, SolutionName, ContextArray, pointSetArray)

    def UpdateAllReports(self):
        self.handle.UpdateAllReports()

    def UpdateReports(self, plotName):
        self.handle.UpdateReports(plotName)

    # BoundarySetup
    # General Commands Recognized by the Boundary/Excitations Module
    def AddAssignmentToBoundary(self, BoundParaArray):
        self.handle.AddAssignmentToBoundary(BoundParaArray)

    def ConvertNportCircuitElementsToPorts(self, NportIDArray):
        self.handle.ConvertNportCircuitElementsToPorts(NportIDArray)

    def CreateNportCircuitElements(self, BoundParaArray):
        self.handle.CreateNportCircuitElements(BoundParaArray)

    def DeleteAllBoundaries(self):
        self.handle.DeleteAllBoundaries()

    def GetBoundaryAssignment(self, BoundaryName):
        return self.handle.GetBoundaryAssignment(BoundaryName)

    def GetBoundaries(self):
        return self.handle.GetBoundaries()

    def GetBoundariesOfType(self, BoundaryType):
        return self.handle.GetBoundariesOfType(BoundaryType)

    def GetDefaultBaseName(self, BoundaryType):
        return self.handle.GetDefaultBaseName(BoundaryType)

    def GetExcitations(self):
        return self.handle.GetExcitations()

    def GetExcitationsOfType(self, ExcitationType):
        return self.handle.GetExcitationsOfType(ExcitationType)

    def GetHybridRegions(self):
        return self.handle.GetHybridRegions()

    def GetHybridRegionsOfType(self, HybridRegionType):
        return self.handle.GetHybridRegionsOfType(HybridRegionType)

    def GetNumBoundaries(self):
        return self.handle.GetNumBoundaries()

    def GetNumBoundariesOfType(self, BoundaryType):
        return self.handle.GetNumBoundariesOfType(BoundaryType)

    def GetNumExcitations(self):
        return self.handle.GetNumExcitations()

    def GetNumExcitationsOfType(self, GetNumExcitationsOfType):
        return self.handle.GetNumExcitationsOfType(GetNumExcitationsOfType)

    def IdentifyNets(self):
        return self.handle.IdentifyNets()

    def GetPortExcitationCounts(self):
        return self.handle.GetPortExcitationCounts()

    def ReassignBoundary(self, BoundParaArray):
        self.handle.ReassignBoundary(BoundParaArray)

    def RemoveAssignmentFromBoundary(self, BoundParaArray):
        self.handle.RemoveAssignmentFromBoundary(BoundParaArray)

    def RenameBoundary(self, OldName, NewName):
        self.handle.RenameBoundary(OldName, NewName)

    def ReprioritizeBoundaries(self, NewOrderArray):
        self.handle.ReprioritizeBoundaries(NewOrderArray)

    def SetDefaultBaseName(self, BoundaryType):
        self.handle.SetDefaultBaseName(BoundaryType)

    # Script Commands for Creating and Modifying Boundaries
    def AssignCircuitPort(self, CircuitPortArray):
        self.handle.AssignCircuitPort(CircuitPortArray)

    def AssignLumpedPort(self, LumpedPortArray):
        self.handle.AssignLumpedPort(LumpedPortArray)

    def AssignPerfectE(self, PerfectEArray):
        """
        Creates a perfect E boundary.
        :param PerfectEArray: [array]
        ["NAME:<BoundName>", "InfGroundPlane:=", True,
         "Objects:=", AssignmentObjects,
         "Faces:=", AssignmentFaces]
        :return: None
        """
        self.handle.AssignPerfectE(PerfectEArray)

    def AssignRadiation(self, RadiationArray):
        """
        Creates a radiation boundary.创建辐射边界条件.
        :param RadiationArray:
            [
                "NAME:Rad1",
                "Objects:="		, ["AirBox"],
                "IsFssReference:="	, False,
                "IsForPML:="		, False
            ]
        :return: None
        """
        self.handle.AssignRadiation(RadiationArray)

    # Script Commands for Creating and Modifying Boundaries in HFSS-IE

    # Script Commands for Creating and Modifying PMLs
    # RadField
    def InsertFarFieldSphereSetup(self, InfSphereParams):
        """
        Creates/inserts a far-field infinite sphere radiation setup.
        :param InfSphereParams:
            [
                "NAME:<SetupName>",
                "UseCustomRadiationSurface:=", <bool>,
                "CustomRadiationSurface:=", <FaceListName>,
                "ThetaStart:=", <value>,
                "ThetaStop:=", <value>,
                "ThetaStep:=", <value>,
                "PhiStart:=", <value>,
                "PhiStop:=", <value>,
                "PhiStep:=", <value>,
                "UseLocalCS:=", <bool>,
                "CoordSystem:=", <CSName>
            ]
        :return: None
        """
        self.handle.InsertFarFieldSphereSetup(InfSphereParams)

    # AnalysisSetup
    def GetSetups(self) -> tuple:
        """

        :return:
        """
        return self.handle.GetSetups()

    def InsertSetup(self, SetupType, AttributesArray):
        """
        :param SetupType:
        :param AttributesArray:
        :return: None
        """
        self.handle.InsertSetup(SetupType, AttributesArray)

    def InsertFrequencySweep(self, SetupName, AttributesArray):
        """

        :param SetupName: "Setup1"

        :param AttributesArray:
            [
                "NAME:Sweep",
                "IsEnabled:="		, True,
                "RangeType:="		, "LinearCount",
                "RangeStart:="		, "8.5GHz",
                "RangeEnd:="		, "12.5GHz",
                "RangeCount:="		, 401,
                "Type:="		, "Interpolating",
                "SaveFields:="		, False,
                "SaveRadFields:="	, False,
                "InterpTolerance:="	, 0.5,
                "InterpMaxSolns:="	, 250,
                "InterpMinSolns:="	, 0,
                "InterpMinSubranges:="	, 1,
                "InterpUseS:="		, True,
                "InterpUsePortImped:="	, False,
                "InterpUsePropConst:="	, True,
                "UseDerivativeConvergence:=", False,
                "InterpDerivTolerance:=", 0.2,
                "UseFullBasis:="	, True,
                "EnforcePassivity:="	, True,
                "PassivityErrorTolerance:=", 0.0001
            ]

            oModule.InsertFrequencySweep(
            "Setup2",
            [
                "NAME:Sweep",
                "IsEnabled:="		, True,
                "RangeType:="		, "LinearCount",
                "RangeStart:="		, "8.5GHz",
                "RangeEnd:="		, "12GHz",
                "RangeCount:="		, 401,
                "Type:="		, "Fast",
                "SaveFields:="		, True,
                "SaveRadFields:="	, False,
                "GenerateFieldsForAllFreqs:=", False
            ])
        :return: None
        """
        self.handle.InsertFrequencySweep(SetupName, AttributesArray)

    def EditFrequencySweep(self, SetupName, SweepName, AttributesArray):
        """
        Modifies an existing frequency sweep. For HFSS-IE use EditSweep [HFSS-IE]
        :param SetupName: str
        :param SweepName: str
        :param AttributesArray: array
        "Setup1", "Sweep",
        [
            "NAME:Sweep",
            "IsEnabled:="		, True,
            "RangeType:="		, "LinearCount",
            "RangeStart:="		, "8.5GHz",
            "RangeEnd:="		, "12.5GHz",
            "RangeCount:="		, 401,
            "Type:="		, "Fast",
            "SaveFields:="		, True,
            "SaveRadFields:="	, False,
            "GenerateFieldsForAllFreqs:=", False
        ]
        :return: None
        """
        self.handle.EditFrequencySweep(SetupName, SweepName, AttributesArray)

    # Solutions
    def ExportNetworkData(self, DesignVariationKey, SolnSelectionArray, FileFormat, OutFile, FreqsArray,
                          DoRenorm, RenormImped, DataType, Pass, ComplexFormat, TouchstoneNumberofDigitsPrecision,
                          IncludeGammaAndImpedanceComments, SupportNon_standardTouchstoneExtensions):
        """
        Exports matrix solution data to a file. Available only for Driven solution types with ports.
        :param DesignVariationKey : 我也不是很清楚这是干嘛用的,看官方文档吧
        :param SolnSelectionArray : Array(<SolnSelector>, <SolnSelector>, ...)
            If more than one array entry, this indicates a combined Interpolating sweep.
            <SolnSelector> : <str>
                Gives solution setup name and solution name, separated by a colon.
        :param FileFormat: int
            Possible values are:
            2 : Tab delimited spreadsheet format (.tab)
            3 : Touchstone (.sNp)
            4 : CitiFile (.cit)
            7 : Matlab (.m)
            8 : Terminal Z0 spreadsheet
        :param OutFile: str
            Full path to the file to write out.
        :param FreqsArray : Array of doubles or "All".
            The frequencies to export. The <FreqsArray> argument contains a vector (e.g. "1GHz","2GHz", ...) to use,
            or "all". To export all frequencies, use Array("all"). If no frequencies are specified, all frequencies are
            used.
        :param DoRenorm: bool
            Specifies whether to renormalize the data before export.
        :param RenormImped: double
            Real impedance value in ohms, for renormalization. Required in syntax, but ignored if DoRenorm is false.
        :param DataType: "S", "Y", or "Z"
            The matrix to export.
        :param Pass: string, from 1 to N.
            The pass to export. This is ignored if the sourceName is a frequency sweep. Leaving out this value or
            specifying -1 gets all passes.
        :param ComplexFormat: "0", "1", or "2"
            The format to use for the exported data.
            0 = Magnitude/Phase.
            1 = Real/Imaginary.
            2 = db/Phase.
        :param TouchstoneNumberofDigitsPrecision: int
            Default if not specified is 15.
        :param IncludeGammaAndImpedanceComments: bool
            Specifies whether to include Gamma and Impedance comments.
        :param SupportNon_standardTouchstoneExtensions: bool
            Specifies whether to support non-standard Touchstone extensions for mixed reference impedances.
        :return: None
        """
        self.handle.ExportNetworkData(DesignVariationKey, SolnSelectionArray, FileFormat, OutFile, FreqsArray,
                                      DoRenorm, RenormImped, DataType, Pass, ComplexFormat,
                                      TouchstoneNumberofDigitsPrecision,
                                      IncludeGammaAndImpedanceComments, SupportNon_standardTouchstoneExtensions)

    # FieldsReporter
    def AddAntennaOverlay(self, OverlayName, SetupName, SolutionName, array, IntrinsicVariations, UseNominalDesign,
                          AntennaParameters):
        self.handle.AddAntennaOverlay(OverlayName, SetupName, SolutionName, array, IntrinsicVariations,
                                      UseNominalDesign, AntennaParameters)


# Conventions Used in this Chapter 本章中使用的约定:
"""
AttributesArray:
Array(  "NAME:Attributes",
        "Name:=", <string>,
        "Flags:=", <string>,
        "Color:=", <string>,
        "Transparency:=", <value>,
        "PartCoordinateSystem:=", <string>,
        "UDMId:=", <string>,
        "MaterialName:=", <string>,
        "SurfaceMaterialName:=", <string>,
        "Solveinside:=", <boolean>,
        "ShellElement:=", <boolean>,
        "ShellElementThickness:=", <string>,
        "IsMaterialEditable:=", <boolean>,
        "UseMaterialAppearance:=", <boolean>,
        "IsLightweight:=", <boolean>)
Flags – Takes a string containing "NonModel" and/or "Wireframe", separated by the # character. 
    For example, "NonModel#Wireframe".
Color – Takes a string containing an RGB triple, formatted as "(<RGB>)". 
    For example, "(255 255 255)".
Transparency – Takes a value between 0 and 1.
PartCoordinateSystem – Orientation of the primitive. The name of one of the defined
coordinate systems should be specified.
UDMId – Takes a string containing an ID.
MaterialName – Takes a string of the material name.
SurfaceMaterialName – Takes a string of the surface material name.
Solveinside – Takes a boolean value.
ShellElement – Takes a boolean value specifying whether or not a shell element is present.
ShellElementThickness – Takes a string containing the shell element thickness. If element
is not present, pass empty string.
IsMaterialEditable – Takes a boolean value.
IsLightweight – Takes a boolean value
"""


# Editor 类的定义 该类对象对应于一个编辑器，如三维建模器、布局或原理图编辑器。此对象用于在编辑器中添加和修改数据。
class Editor(Object):
    def __init__(self):
        super(Editor, self).__init__()
        self.handle = None

    # 3D Modeler
    # Draw Menu Commands
    def Create3DComponent(self, GeometryData, DesignData, FileName, ImageFile):
        """
        Creates a 3D component.
        :param GeometryData: [array], Structured array containing geometry:
            Array("NAME:GeometryData","ComponentName:=","<string>","Owner:=", "<string>","Email:=", "<string>",
            "Company:=", "<string>", "Version:=", "<string>", "Date:=", "<string>", "Notes:=", "<string>",
            "HasLabel:=", <boolean>, "IncludedParts:=", <array>, "IncludedCS:=", <array>, "ReferenceCS:=", <string>,
            "IncludedParameters:=", <array>, "ParameterDescription:=", <array>)
        :param DesignData: [array] Structured array containing design data:
            Array("NAME:DesignData", "Boundaries:=", <array>, "Excitations:=", <array>, "MeshOperations:=", <array>)
        :param FileName: [str]
            Full file path to 3D component.
        :param ImageFile: [array]
            Optional. Structured array containing file path of 3D component image:
            Array("NAME:ImageFile", "ImageFile:=", <string>)
        :return: None
        """
        self.handle.Create3DComponent(GeometryData, DesignData, FileName, ImageFile)

    def CreateBondwire(self, Parameters, Attributes):
        """
        Creates a bondwire.
        :param Parameters: [array]
            Array("NAME:BondwireParameters","WireType:=", <string("JEDEC_4Points", "JEDEC_5Points", or "LOW")>,
            "WireDiameter:=", <string>, "NumSides:=", <value>, "XPadPos:=", <value>, "YPadPos:=", <value>,
            "ZPadPos:=", <value>, "XDir:=", <value>, "YDir:=", <value>, "ZDir:=", <value>, "Distance:=", <value>,
            "h1:=", <value>, "h2:=", <value>, "alpha:=", <value>, "beta:=", <value>,
            "WhichAxis:=", <string("X","Y", or "Z")>, "ReverseDirection:=", <boolean>)
        for example:
        ["NAME:BondwireParameters", "WireType:=" , "JEDEC_4Points", "WireDiameter:=" , "0.025mm", "NumSides:=" , "6",
         "XPadPos:=" , "1.6mm", "YPadPos:=" , "-0.2mm", "ZPadPos:=" , "0mm", "XDir:=" , "-2.2mm", "YDir:=" , "-1.4mm",
         "ZDir:=" , "0mm", "Distance:=" , "2.60768096208106mm", "h1:=" , "0.2mm", "h2:=" , "0mm", "alpha:=" , "80deg",
         "beta:=" , "0", "WhichAxis:=" , "Z", "ReverseDirection:=" , True ],
        :param Attributes: [array] Structured array. See: AttributesArray.
        :return: None
        """
        self.handle.CreateBondwire(Parameters, Attributes)

    def CreateBox(self, Parameters, Attributes):
        """
        Creates a box.
        :param Parameters: Structured array.
            Array("NAME:BoxParameters", "XPosition:=", <string>, "YPosition:=", <string>,
            "ZPosition:=", <string>, "XSize:=", <string>, "YSize:=", <string>, "ZSize:=", <string>)
        :param Attributes: [array] Structured array. See: AttributesArray.
        :return: None
        """
        self.handle.CreateBox(Parameters, Attributes)

    def CreateCircle(self, Parameters, Attributes):
        """
        Creates a circle.
        :param Parameters: Structured array.
            Array("NAME:CircleParameters", "IsCovered:=", <boolean>, "XCenter:=", <value>, "YCenter:=", <value>,
            "ZCenter:=", <value>, "Radius:=", <value>, "WhichAxis:=", <string> "NumSegments:=",
            <string containing integer>)
        :param Attributes: [array] Structured array. See: AttributesArray.
        :return: None
        """
        self.handle.CreateCircle(Parameters, Attributes)

    def CreateCone(self, Parameters, Attributes):
        """

        :param Parameters: [array] Structured array.
        :param Attributes: [array] Structured array. See: AttributesArray.
        :return:
        """
        self.handle.CreateCone(Parameters, Attributes)

    def CreateCutplane(self, Parameters, Attributes):
        self.handle.CreateCutplane(Parameters, Attributes)

    def CreateCylinder(self, Parameters, Attributes):
        self.handle.CreateCylinder(Parameters, Attributes)

    def CreateEllipse(self, Parameters, Attributes):
        self.handle.CreateEllipse(Parameters, Attributes)

    def CreateEquationCurve(self, Parameters, Attributes):
        self.handle.CreateEquationCurve(Parameters, Attributes)

    def CreateHelix(self, Parameters, Attributes):
        self.handle.CreateHelix(Parameters, Attributes)

    def CreatePoint(self, Parameters, Attributes):
        self.handle.CreatePoint(Parameters, Attributes)

    def CreatePolyline(self, Parameters, Attributes):
        self.handle.CreatePolyline(Parameters, Attributes)

    def CreateRectangle(self, Parameters, Attributes):
        """
        创建一个矩形.

        :param Parameters: [Structured Array]
            Array("NAME:RectangleParameters", "IsCovered:=" , boolean, "XStart:=" , string, "YStart:=" , string,
                  "ZStart:=" , string, "Width:=" , string, "Height:=" , string,
                  "WhichAxis:=" , string "X", "Y", or "Z")

        :param Attributes: [Structured Array]  See: AttributesArray.
        :return: None
        """
        self.handle.CreateRectangle(Parameters, Attributes)

    def CreateRegion(self, Parameters, Attributes):
        """
        Creates a region containing the design.
        :param Parameters:
        [
            "NAME:RegionParameters",
            "+XPaddingType:="	, "Percentage Offset", "+XPadding:=", "25mm",
            "-XPaddingType:="	, "Percentage Offset", "-XPadding:=", "25mm",
            "+YPaddingType:="	, "Percentage Offset", "+YPadding:=", "25mm",
            "-YPaddingType:="	, "Percentage Offset", "-YPadding:=", "25mm",
            "+ZPaddingType:="	, "Percentage Offset", "+ZPadding:=", "25mm",
            "-ZPaddingType:="	, "Percentage Offset", "-ZPadding:=", "25mm"
        ],
        :param Attributes:
         [
            "NAME:Attributes",
            "Name:="		, "Region",
            "Flags:="		, "Wireframe#",
            "Color:="		, "(143 175 143)",
            "Transparency:="	, 0,
            "PartCoordinateSystem:=", "Global",
            "UDMId:="		, "",
            "MaterialValue:="	, "\"vacuum\"",
            "SurfaceMaterialValue:=", "\"\"",
            "SolveInside:="		, True,
            "ShellElement:="	, False,
            "ShellElementThickness:=", "nan ",
            "IsMaterialEditable:="	, True,
            "UseMaterialAppearance:=", False,
            "IsLightweight:="	, False
        ]
        :return:
        """
        self.handle.CreateRegion(Parameters, Attributes)

    def CreateRegularPolygon(self, Parameters, Attributes):
        self.handle.CreateRegularPolygon(Parameters, Attributes)

    def CreateRegularPolyhedron(self, Parameters, Attributes):
        self.handle.CreateRegularPolyhedron(Parameters, Attributes)

    def CreateSphere(self, Parameters, Attributes):
        self.handle.CreateSphere(Parameters, Attributes)

    def CreateSpiral(self, Parameters, Attributes):
        self.handle.CreateSpiral(Parameters, Attributes)

    def CreateTorus(self, Parameters, Attributes):
        self.handle.CreateTorus(Parameters, Attributes)

    def CreateUserDefinedModel(self, Parameters):
        self.handle.CreateUserDefinedModel(Parameters)

    def CreateUserDefinedPart(self, Parameters, Attributes):
        self.handle.CreateUserDefinedPart(Parameters, Attributes)

    def Edit3DComponent(self, compName, Parameters):
        self.handle.Edit3DComponent(compName, Parameters)

    def EditPolyline(self, SelectionsArray, Parameters):
        self.handle.EditPolyline(SelectionsArray, Parameters)

    def Get3DComponentDefinitionNames(self):
        return self.handle.Get3DComponentDefinitionNames()

    def Get3DComponentInstanceNames(self, DefinitionName):
        return self.handle.Get3DComponentInstanceNames(DefinitionName)

    def Get3DComponentMaterialNames(self, InstanceName):
        return self.handle.Get3DComponentMaterialNames(InstanceName)

    def Get3DComponentMaterialProperties(self, MaterialName):
        return self.handle.Get3DComponentMaterialProperties(MaterialName)

    def Get3DComponentParameters(self, ComponentName):
        return self.handle.Get3DComponentParameters(ComponentName)

    def Insert3DComponent(self, ComponentData):
        self.handle.Insert3DComponent(ComponentData)

    def InsertPolylineSegment(self, Parameters):
        self.handle.InsertPolylineSegment(Parameters)

    def SweepAlongPath(self, SelectionsArray, PathSweepParametersArray):
        self.handle.SweepAlongPath(SelectionsArray, PathSweepParametersArray)

    def SweepAlongVector(self, SelectionsArray, VecSweepParametersArray):
        self.handle.SweepAlongVector(SelectionsArray, VecSweepParametersArray)

    def SweepAroundAxis(self, SelectionsArray, AxisSweepParametersArray):
        self.handle.SweepAroundAxis(SelectionsArray, AxisSweepParametersArray)

    def SweepFacesAlongNormal(self, SelectionsArray, Parameters):
        self.handle.SweepFacesAlongNormal(SelectionsArray, Parameters)

    def SweepFacesAlongNormalWithAttributes(self, SelectionsArray, Parameters, AttributesArray):
        self.handle.SweepFacesAlongNormalWithAttributes(SelectionsArray, Parameters, AttributesArray)

    def UpdateComponentDefinition(self, data):
        self.handle.UpdateComponentDefinition(data)

    # Edit Menu Commands
    def Copy(self, SelectionsArray):
        self.handle.Copy(SelectionsArray)

    def DeletePolylinePoint(self, DeletePointArray):
        self.handle.DeletePolylinePoint(DeletePointArray)

    def DuplicateAlongLine(self, SelectionsArray, ParametersArray, OptionsArray, CreateGroup):
        self.handle.DuplicateAlongLine(SelectionsArray, ParametersArray, OptionsArray, CreateGroup)

    def DuplicateAroundAxis(self, SelectionsArray, ParametersArray, OptionsArray, CreateGroup):
        self.handle.DuplicateAroundAxis(SelectionsArray, ParametersArray, OptionsArray, CreateGroup)

    def DuplicateMirror(self, SelectionsArray, ParametersArray, OptionsArray, CreateGroup):
        self.handle.DuplicateMirror(SelectionsArray, ParametersArray, OptionsArray, CreateGroup)

    def Mirror(self, SelectionsArray, MirrorParameters):
        self.handle.Mirror(SelectionsArray, MirrorParameters)

    def Move(self, SelectionsArray, TranslateParameters):
        self.handle.Move(SelectionsArray, TranslateParameters)

    def OffsetFaces(self, SelectionsArray, OffsetParameters):
        self.handle.OffsetFaces(SelectionsArray, OffsetParameters)

    def Paste(self):
        self.handle.Paste()

    def Rotate(self, SelectionsArray, RotateParameters):
        self.handle.Rotate(SelectionsArray, RotateParameters)

    def Scale(self, SelectionsArray, ScaleParameters):
        self.handle.Scale(SelectionsArray, ScaleParameters)

    # Modeler Menu Commands
    def AssignMaterial(self, SelectionsArray, AttributesArray):
        self.handle.AssignMaterial(SelectionsArray, AttributesArray)

    def Chamfer(self, SelectionsArray, Parameters):
        self.handle.Chamfer(SelectionsArray, Parameters)

    def Connect(self, SelectionsArray):
        self.handle.Connect(SelectionsArray)

    def CoverLines(self, SelectionsArray):
        self.handle.CoverLines(SelectionsArray)

    def CoverSurfaces(self, SelectionsArray):
        self.handle.CoverSurfaces(SelectionsArray)

    def CreateEntityList(self, Parameters, AttributesArray):
        self.handle.CreateEntityList(Parameters, AttributesArray)

    def CreateFaceCS(self, Parameters, AttributesArray):
        self.handle.CreateFaceCS(Parameters, AttributesArray)

    def CreateGroup(self, Parameters):
        self.handle.CreateGroup(Parameters)

    def CreateObjectCS(self, Parameters, AttributesArray):
        self.handle.CreateObjectCS(Parameters, AttributesArray)

    def CreateObjectFromEdges(self, SelectionsArray, ParametersArray, CreateGroupsForNewObjects):
        self.handle.CreateObjectFromEdges(SelectionsArray, ParametersArray, CreateGroupsForNewObjects)

    def CreateObjectFromFaces(self, SelectionsArray, Parameters, CreateGroupsForNewObjects):
        self.handle.CreateObjectFromFaces(SelectionsArray, Parameters, CreateGroupsForNewObjects)

    def CreateRelativeCS(self, Parameters, AttributesArray):
        self.handle.CreateRelativeCS(Parameters, AttributesArray)

    def DeleteEmptyGroups(self, Parameters):
        self.handle.DeleteEmptyGroups(Parameters)

    def DeleteLastOperation(self, SelectionsArray):
        self.handle.DeleteLastOperation(SelectionsArray)

    def DetachFaces(self, SelectionsArray, Parameters, DetachFacesArray):
        self.handle.DetachFaces(SelectionsArray, Parameters, DetachFacesArray)

    def EditEntityList(self, SelectionsArray, Parameters):
        self.handle.EditEntityList(SelectionsArray, Parameters)

    def EditFaceCS(self, Parameters, AttributesArray):
        self.handle.EditFaceCS(Parameters, AttributesArray)

    def EditObjectCS(self, Parameters, AttributesArray):
        self.handle.EditObjectCS(Parameters, AttributesArray)

    def EditRelativeCS(self, Parameters, AttributesArray):
        self.handle.EditRelativeCS(Parameters, AttributesArray)

    def Export(self, Parameters):
        self.handle.Export(Parameters)

    def ExportModelImageToFile(self, path, width, height, Parameters):
        self.handle.ExportModelImageToFile(path, width, height, Parameters)

    def Fillet(self, SelectionsArray, Parameters):
        self.handle.Fillet(SelectionsArray, Parameters)

    def FlattenGroup(self, GroupID):
        self.handle.FlattenGroup(GroupID)

    def GenerateHistory(self, SelectionsArray):
        self.handle.GenerateHistory(SelectionsArray)

    def GetActiveCoordinateSystem(self):
        return self.handle.GetActiveCoordinateSystem()

    def GetCoordinateSystems(self):
        return self.handle.GetCoordinateSystems()

    def HealObject(self, SelectionsArray, Parameters):
        self.handle.HealObject(SelectionsArray, Parameters)

    def Import(self, Parameters):
        self.handle.Import(Parameters)

    def ImportDXF(self, Parameters):
        self.handle.ImportDXF(Parameters)

    def ImportGDSII(self, Parameters):
        self.handle.ImportGDSII(Parameters)

    def Intersect(self, SelectionsArray, Parameters):
        self.handle.Intersect(SelectionsArray, Parameters)

    def MoveCStoEnd(self, SelectionsArray):
        self.handle.MoveCStoEnd(SelectionsArray)

    def MoveEntityToGroup(self, Objects, MoveEntityToGroup):
        self.handle.MoveEntityToGroup(Objects, MoveEntityToGroup)

    def MoveFaces(self, SelectionsArray, MoveFacesParametersArray):
        self.handle.MoveFaces(SelectionsArray, MoveFacesParametersArray)

    def ProjectSheet(self, SelectionsArray, Parameters):
        self.handle.ProjectSheet(SelectionsArray, Parameters)

    def PurgeHistory(self, SelectionsArray):
        self.handle.PurgeHistory(SelectionsArray)

    def ReplaceWith3DComponent(self, GeometryData, DesignData, ImageFile):
        self.handle.ReplaceWith3DComponent(GeometryData, DesignData, ImageFile)

    def Section(self, SelectionsArray, Parameters):
        self.handle.Section(SelectionsArray, Parameters)

    def SeparateBody(self, SelectionsArray):
        self.handle.SeparateBody(SelectionsArray)

    def SetModelUnits(self, Parameters):
        """
        Sets the model units.
        :param Parameters: Array
            Array("NAME:UnitsParameter","Units:=",<string>,"Rescale:=",<boolean True to rescale model; else False>)
        :return: None
        """
        self.handle.SetModelUnits(Parameters)

    def SetWCS(self, Parameters):
        self.handle.SetWCS(Parameters)

    def ShowWindow(self):
        self.handle.ShowWindow()

    def Split(self, SelectionsArray, Parameters):
        self.handle.Split(SelectionsArray, Parameters)

    def Subtract(self, Selections, Parameters):
        self.handle.Subtract(Selections, Parameters)

    # def SweepFacesAlongNormal(self, SelectionsArray, Parameters):
    #     self.handle.SweepFacesAlongNormal(SelectionsArray, Parameters)

    def ThickenSheet(self, SelectionsArray, Parameters):
        self.handle.ThickenSheet(SelectionsArray, Parameters)

    def UncoverFaces(self, SelectionsArray, Parameters):
        self.handle.UncoverFaces(SelectionsArray, Parameters)

    def Unite(self, SelectionsArray, Parameters):
        self.handle.Unite(SelectionsArray, Parameters)

    def Ungroup(self, Groups):
        self.handle.Ungroup(Groups)

    def WrapSheet(self, SelectionsArray, Parameters):
        self.handle.WrapSheet(SelectionsArray, Parameters)

    # Other oEditor Commands
    def AddViewOrientation(self, Name, Global, Orientation, ProjectionParams):
        self.handle.AddViewOrientation(Name, Global, Orientation, ProjectionParams)

    def BreakUDMConnection(self, SelectionsArray):
        self.handle.BreakUDMConnection(SelectionsArray)

    # def ChangeProperty(self, args):
    #     self.handle.ChangeProperty(args)

    def Delete(self, SelectionsArray):
        self.handle.Delete(SelectionsArray)

    def FitAll(self):
        self.handle.FitAll()

    def GetBodyNamesByPosition(self, positionParameters):
        return self.handle.GetBodyNamesByPosition(positionParameters)

    def GetEdgeByPosition(self, positionParameters):
        return self.handle.GetEdgeByPosition(positionParameters)

    def GetEdgeIDsFromFace(self, faceID):
        return self.handle.GetEdgeIDsFromFace(faceID)

    def GetEdgeIDsFromObject(self, objectName):
        return self.handle.GetEdgeIDsFromObject(objectName)

    def GetEntityListIDByName(self, listName):
        return self.handle.GetEntityListIDByName(listName)

    def GetFaceArea(self, faceName):
        return self.handle.GetFaceArea(faceName)

    def GetFaceCenter(self, planarFaceID):
        return self.handle.GetFaceCenter(planarFaceID)

    def GetFaceByPosition(self, FaceParameters):
        return self.handle.GetFaceByPosition(FaceParameters)

    def GetFaceIDs(self, objectName):
        return int(self.handle.GetFaceIDs(objectName)[0])

    def GetGeometryModelerMode(self):
        return self.handle.GetGeometryModelerMode()

    def GetMatchedObjectName(self, wildcardText):
        return self.handle.GetMatchedObjectName(wildcardText)

    def GetModelBoundingBox(self):
        return self.handle.GetModelBoundingBox()

    def GetModelUnits(self):
        return self.handle.GetModelUnits()

    def GetNumObjects(self):
        return self.handle.GetNumObjects()

    def GetObjectIDByName(self, objectName):
        return self.handle.GetObjectIDByName(objectName)

    def GetObjectName(self):
        return self.handle.GetObjectIDByName()

    def GetObjectNameByFaceID(self, faceID):
        return self.handle.GetObjectNameByFaceID(faceID)

    def GetObjectsByMaterial(self, material):
        return self.handle.GetObjectsByMaterial(material)

    def GetObjectsInGroup(self, group):
        return self.handle.GetObjectsInGroup(group)

    def GetPropNames(self, IncludeReadOnly):
        return self.handle.GetPropNames(IncludeReadOnly)

    def GetPropValue(self, propPath):
        return self.handle.GetPropValue(propPath)

    def GetSelections(self):
        return self.handle.GetSelections()

    def GetUserPosition(self, prompt):
        return self.handle.GetUserPosition(prompt)

    def GetVertexIDsFromEdge(self, edgeID):
        return self.handle.GetVertexIDsFromEdge(edgeID)

    def GetVertexIDsFromFace(self, faceID):
        return self.handle.GetVertexIDsFromFace(faceID)

    def GetVertexIDsFromObject(self, objectName):
        return self.handle.GetVertexIDsFromObject(objectName)

    def GetVertexPosition(self, vertexID):
        return self.handle.GetVertexPosition(vertexID)

    def OpenExternalEditor(self, SelectionsArray):
        self.handle.OpenExternalEditor(SelectionsArray)

    def PageSetup(self, Parameters):
        self.handle.PageSetup(Parameters)

    def RenamePart(self, renameParametersArray):
        self.handle.RenamePart(renameParametersArray)

    def SetPropValue(self, propPath, value):
        return self.handle.SetPropValue(propPath, value)


# AnsoftApp 类的定义 该类对象为IronPython或 VBScript 提供了一个访问 Ansoft.ElectronicsDesktop 产品。
class AnsoftApp(object):
    def __init__(self, hfss_app):
        self.handle = win32com.client.Dispatch(hfss_app)
        self.oDesktop = None
        pass

    def GetAppDesktop(self):
        return self.handle.GetAppDesktop()
        pass

    def OpenAnsoftAPP(self):
        # del self.handle
        os.system("F:\\AnsysEM\\AnsysEM21.1\\Win64\\reg_ansysedt.exe  -shortcut")
        # self.handle = win32com.client.Dispatch(hfssAppName)
        # self.handle = win32com.client.Dispatch(hfssAppName)

        pass


# 下面为 hfss 各对象的实例化
# hfss 脚本控制接口
hfssAppName = "AnsoftHfss.HfssScriptInterface"

# 各对象实例化
oAnsoftApp = AnsoftApp(hfssAppName)
oDesktop = Desktop(oAnsoftApp.handle)
# 下面对象句柄初始化为空 使用时应先对句柄进行相应的初始化
oProject = Project()
oDesign = Design()
oEditor = Editor()
oModule = Module()

# 测试
if __name__ == '__main__':
    oDesktop.NewProject()
    oProject.handle = oDesktop.GetActiveProject()
    oDesktop.CloseProject('Project4')

    time.sleep(20)  # 等待 20 秒后关闭程序
    oDesktop.QuitApplication()  # 退出桌面程序
    del oProject
    del oDesktop
    del oAnsoftApp
