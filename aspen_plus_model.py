import pywinauto.application
import win32com.client
from typing import Literal

import time
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, LogLocator, AutoLocator
import pandas as pd
import numpy as np

def startup_aspen_plus(filename, visible=True, suppress_dialogs=True):
    Aspen = win32com.client.Dispatch('Apwn.Document')
    Aspen.InitFromFile2(filename)
    Aspen.Visible = int(visible)  # Convert boolean to int for COM compatibility
    Aspen.SuppressDialogs = suppress_dialogs

    app = pywinauto.application.Application(backend="uia").connect(process=Aspen.ProcessId)
    asp_dialog = app["Dialog"]

    return Aspen, asp_dialog

def kill_aspen_plus():
    # kill all Aspen Plus files
    import subprocess
    from subprocess import CalledProcessError
    
    try:
        subprocess.run(["taskkill", "/f", "/im", "AspenPlus.exe"], check=True)
    except CalledProcessError:
        print('Aspen Plus has already been killed. ')

class Tool:
    def __init__(self, Aspen, Asp):
        self.Aspen = Aspen
        self.Asp = Asp

    def cd_Properties(self):
        Asp = self.Asp
        Asp.children()[3].children()[0].children()[6].select()

    def cd_Simulation(self):
        Asp = self.Asp
        Asp.children()[3].children()[0].children()[7].select()

    # under Properties
    def click_Components(self):
        Asp = self.Asp
        Asp.child_window(title="Components", auto_id="igRibbon_btnProperties_Components", control_type="Button").click()

    def click_Review(self):
        Asp = self.Asp
        Asp.child_window(title="Review", auto_id="PART_BUTTON", control_type="Button").click()

    def Run_prop(self):
        Asp = self.Asp
        tool_bar = Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", class_name='QuickAccessToolbar', control_type="ToolBar")
        run = tool_bar.child_window(title="Run", auto_id='igRibbon_btnRunProp', control_type="Button")
        if run.legacy_properties()['State'] != 1:
            run.click()

    def next_prop(self):
        Asp = self.Asp
        tool_bar = Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", class_name='QuickAccessToolbar', control_type="ToolBar")
        next_prop = tool_bar.child_window(title="Next", auto_id='igRibbon_btnPropertiesNextInput', control_type="Button")
        if next_prop.legacy_properties()['State'] != 1:
            next_prop.click()

    def results_available(self): # return True or False
        Asp = self.Asp
        results_available = Asp.child_window(title="Results Available", auto_id="PART_StatusText", class_name="TextBlock").exists()
        results_available_with_errors = Asp.child_window(title="Results Available with Errors", auto_id="PART_StatusText", class_name="TextBlock").exists()
        return results_available

    def input_complete(self):
        Asp = self.Asp
        required_properties_input_complete = Asp.child_window(title="Required Properties Input Complete", auto_id="PART_StatusText", class_name="TextBlock").exists()
        required_input_incomplete = Asp.child_window(title="Required Input Incomplete", auto_id="PART_StatusText", class_name="TextBlock").exists()
        if required_properties_input_complete:
            return True
        elif required_input_incomplete:
            return False

    def make_complete(self):
        Asp = self.Asp
        
        while not self.input_complete(): # click next until input complete or other problems happened
            completion_window = Asp.child_window(title="Completion Status", auto_id="Window_1", class_name="Window")
            self.next_prop()
            
            if completion_window.exists() or self.results_available() or self.input_complete():
                break

    # for setup 
    def set_base_method(self, method='ELECNRTL', eos='ESPRV2'):
        Aspen = self.Aspen
        Aspen.Application.Tree.FindNode(r"\Data\Properties\Specifications\Input\GBASEOPSET").Value = method

        ### other settings
        Aspen.Application.Tree.FindNode(r"\Data\Properties\Specifications\Input\GMODIFY").Value = 'YES'
        Aspen.Application.Tree.FindNode(r"\Data\Properties\Specifications\Input\GVAPOREOS").Value = eos
        Aspen.Application.Tree.FindNode(f"\\Data\\Properties\\Property Methods\\{method}\\Input\\MODELNAME\\#4").Value = 'ESPRV20' # this fixes the bug in Aspen
        
        # add Chemistry
        # Aspen.Application.Tree.FindNode(r"\Data\Properties\Specifications\Input\GCHEMISTRY").Value = 'YUE-CHE'

    def insert_component(self, alias): # insert the component which exists in Aspen Database
        Aspen = self.Aspen
        Aspen.Application.Tree.Data.Components.Specifications.Input.TYPE.Elements.InsertRow(0,0)
        Aspen.Application.Tree.Data.Components.Specifications.Input.TYPE.Elements.LabelNode(0,0)[0].Value = alias

    def set_scalar_property(self, alias, property, value, unit=None): # unit=None -> no unit or default unit
        Aspen = self.Aspen
        review_path = '\\Data\\Properties\\Parameters\\Pure Components\\REVIEW-1\\Input\\VALUE'
        review_unit_path = "\\Data\\Properties\\Parameters\\Pure Components\\REVIEW-1\\Input\\UNITLABEL"
        
        self.check_review_node()
        
        try:
            Aspen.Application.Tree.FindNode(f"{review_path}\\{property}\{alias}").Value = value
            if unit:
                Aspen.Application.Tree.FindNode(f"{review_unit_path}\\{property}").Value = unit
        except AttributeError:
            print('There is no such alias/property/unit.')

    def check_review_node(self): # check if review has node, if no, grow it
        Aspen = self.Aspen
        try:
            review_path = '\\Data\\Properties\\Parameters\\Pure Components\\REVIEW-1\\Input\\VALUE'
            Aspen.Application.Tree.FindNode(review_path).Elements
        except AttributeError:
            self.click_Components()
            self.click_Review()
            # sleep for growing node
            time.sleep(1)

    def comp_list(self):
        comp_list = []
        i = 0
        comp_path = "\\Data\\Components\\Comp-Lists\\GLOBAL\\Input\\CID"
        while True:
            try:
                comp_list.append(self.Aspen.Application.Tree.FindNode(f'{comp_path}\\#{i}').Value)
                i += 1
            except AttributeError:
                break
        return comp_list

class FluidPropModel:
    def __init__(self, Aspen, Asp, tool):
        self.Aspen = Aspen
        self.Asp = Asp
        self.tool = tool

    # pure, scalar
    def set_fluid_properties(self, alias, h, g, mw, tc, pc, vc, zc, omega):
        # need click_Review first
        # setup ion's H, G, charge, mw, type(for Cp), Pvap(nonvolatile) unit: kcal/mol
        Aspen = self.Aspen
        alias = alias.upper()
        
        review_path = '\\Data\\Properties\\Parameters\\Pure Components\\REVIEW-1\\Input\\VALUE'
        review_unit_path = "\\Data\\Properties\\Parameters\\Pure Components\\REVIEW-1\\Input\\UNITLABEL"
        
        # check if review has node
        self.tool.check_review_node()
        
        # change units to TEAM style
        Aspen.Application.Tree.FindNode(f"{review_unit_path}\\TC").Value = "K"
        Aspen.Application.Tree.FindNode(f"{review_unit_path}\\PC").Value = "bar"
        Aspen.Application.Tree.FindNode(f"{review_unit_path}\\VC").Value = "cum/kmol"
        Aspen.Application.Tree.FindNode(f"{review_unit_path}\\DGFORM").Value = "kcal/mol"
        Aspen.Application.Tree.FindNode(f"{review_unit_path}\\DHFORM").Value = "kcal/mol"
        
        def insert_new_component(alias): # at Components
            Aspen.Application.Tree.Data.Components.Specifications.Input.TYPE.Elements.InsertRow(0,0)
            Aspen.Application.Tree.Data.Components.Specifications.Input.TYPE.Elements.LabelNode(0,0)[0].Value = alias
            Aspen.Application.Tree.FindNode(f"\\Data\\Components\\Specifications\\Input\\ANAME1\\{alias}").Value = ''
        
        def setup_fluid_thermo_properties(alias, h, g, mw, tc, pc, vc, zc, omega):
            Aspen.Application.Tree.FindNode(review_path).Elements.InsertRow(1,0)
            Aspen.Application.Tree.FindNode(review_path).Elements.LabelNode(1,0)[0].Value = alias
            Aspen.Application.Tree.FindNode(f"{review_path}\\DHAQFM\{alias}").Value = h
            Aspen.Application.Tree.FindNode(f"{review_path}\\DGAQFM\{alias}").Value = g
            Aspen.Application.Tree.FindNode(f"{review_path}\\MW\{alias}").Value = mw
            Aspen.Application.Tree.FindNode(f"{review_path}\\TC\{alias}").Value = tc
            Aspen.Application.Tree.FindNode(f"{review_path}\\PC\{alias}").Value = pc
            Aspen.Application.Tree.FindNode(f"{review_path}\\VC\{alias}").Value = vc
            Aspen.Application.Tree.FindNode(f"{review_path}\\ZC\{alias}").Value = zc
            Aspen.Application.Tree.FindNode(f"{review_path}\\OMEGA\{alias}").Value = omega
            
        insert_new_component(alias)
        setup_fluid_thermo_properties(alias, h, g, mw, mw, tc, pc, vc, zc, omega)

    # pure, T-dependent
    def set_PLXANT(self, comp, val_list):
        Aspen = self.Aspen
        plxant_path = "Data\\Properties\\Parameters\\Pure Components\\PLXANT-1\\Input"
        pre_unit_path = f"{plxant_path}\\UNITLABEL2\\PLXANT\\{comp}"
        temp_unit_path = f"{plxant_path}\\TUNITLABEL2\\PLXANT\\{comp}"

        Aspen.Application.Tree.FindNode(f"{plxant_path}\\VAL1\\PLXANT").Elements.InsertRow(0,0)
        Aspen.Application.Tree.FindNode(f"{plxant_path}\\VAL1\\PLXANT").Elements.LabelNode(0,0)[0].Value = comp

        Aspen.Application.Tree.FindNode(temp_unit_path).Value = "K"
        Aspen.Application.Tree.FindNode(pre_unit_path).Value = "Pa"

        for i, val in enumerate(val_list):
            val_path = f"{plxant_path}\\VAL{str(i+1)}\\PLXANT\\{comp}"
            Aspen.Application.Tree.FindNode(val_path).Value = val

    def set_CPIGDP(self, comp, val_list):
        Aspen = self.Aspen
        cpigdp_path = "Data\\Properties\\Parameters\\Pure Components\\CPIGDP-1\\Input"
        cp_unit_path = f"{cpigdp_path}\\UNITLABEL2\\CPIGDP\\{comp}"
        temp_unit_path = f"{cpigdp_path}\\TUNITLABEL2\\CPIGDP\\{comp}"

        Aspen.Application.Tree.FindNode(f"{cpigdp_path}\\VAL1\\CPIGDP").Elements.InsertRow(0,0)
        Aspen.Application.Tree.FindNode(f"{cpigdp_path}\\VAL1\\CPIGDP").Elements.LabelNode(0,0)[0].Value = comp

        Aspen.Application.Tree.FindNode(temp_unit_path).Value = "K"
        Aspen.Application.Tree.FindNode(cp_unit_path).Value = "cal/mol-K"

        for i, val in enumerate(val_list):
            val_path = f"{cpigdp_path}\\VAL{str(i+1)}\\CPIGDP\\{comp}"
            Aspen.Application.Tree.FindNode(val_path).Value = val

    def set_DHVLDP(self, comp, val_list):
        Aspen = self.Aspen
        dhvldp_path = "Data\\Properties\\Parameters\\Pure Components\\DHVLDP-1\\Input"
        cp_unit_path = f"{dhvldp_path}\\UNITLABEL2\\DHVLDP\\{comp}"
        temp_unit_path = f"{dhvldp_path}\\TUNITLABEL2\\DHVLDP\\{comp}"

        Aspen.Application.Tree.FindNode(f"{dhvldp_path}\\VAL1\\DHVLDP").Elements.InsertRow(0,0)
        Aspen.Application.Tree.FindNode(f"{dhvldp_path}\\VAL1\\DHVLDP").Elements.LabelNode(0,0)[0].Value = comp

        Aspen.Application.Tree.FindNode(temp_unit_path).Value = "K"
        Aspen.Application.Tree.FindNode(cp_unit_path).Value = "cal/mol"

        for i, val in enumerate(val_list):
            val_path = f"{dhvldp_path}\\VAL{str(i+1)}\\DHVLDP\\{comp}"
            Aspen.Application.Tree.FindNode(val_path).Value = val

    # binary
    def clear_henry(self):
        Aspen = self.Aspen
        henry_path = '\\Data\\Properties\\Parameters\\Binary Interaction\\HENRY-1\\Input'
        Aspen.Application.Tree.FindNode(f"{henry_path}\\VAL1\\HENRY").RemoveAll()

    def set_henry_comp(self, gas):
        Aspen = self.Aspen
        hen_id = 'YUE-HEN'
        id_path = f'\\Data\\Components\\Henry-Comps\\{hen_id}\\Input\\CID'
        
        Aspen.Application.Tree.FindNode(r"\Data\Properties\Specifications\Input\GHEN").Value = hen_id
        Aspen.Application.Tree.FindNode(id_path).Elements.InsertRow(0,0)
        Aspen.Application.Tree.FindNode(f"{id_path}\\#0").Value = gas
        

    def set_henry(self, gas, solvent, val_list: list=[], temp_unit='K', pres_unit='bar'): # [A12, B12, C12, D12, T_l, T_u]
        Aspen = self.Aspen
        # hen_id = 'YUE-HEN'
        # id_path = f'\\Data\\Components\\Henry-Comps\\{hen_id}\\Input\\CID'
        
        # # check if gas is in Henry Comps, if not, add it
        # henry_components = []
        # index = 0
        # try:
        #     while True:
        #         henry_components.append(Aspen.Application.Tree.FindNode(f"{id_path}\\#{str(index)}").Value)
        #         index += 1
                
        # except AttributeError:
        #     if gas in henry_components:
        #         Aspen.Application.Tree.FindNode(r"\Data\Properties\Specifications\Input\GHEN").Value = hen_id
        #         Aspen.Application.Tree.FindNode(id_path).Elements.InsertRow(0,0)
        #         Aspen.Application.Tree.FindNode(f"{id_path}\\#0").Value = gas
        
        henry_path = '\\Data\\Properties\\Parameters\\Binary Interaction\\HENRY-1\\Input'
        Aspen.Application.Tree.FindNode(f"{henry_path}\\VAL1\\HENRY").Elements.InsertRow(0,0)
        Aspen.Application.Tree.FindNode(f"{henry_path}\\VAL1\\HENRY").Elements.LabelNode(0,0)[0].Value = gas
        Aspen.Application.Tree.FindNode(f"{henry_path}\\VAL1\\HENRY").Elements.LabelNode(1,0)[0].Value = solvent
        Aspen.Application.Tree.FindNode(f'{henry_path}\\TUNITLABEL2\\HENRY\\{gas}\\{solvent}').Value = temp_unit
        Aspen.Application.Tree.FindNode(f'{henry_path}\\UNITLABEL2\\HENRY\\{gas}\\{solvent}').Value = pres_unit
        
        for i, val in enumerate(val_list):
            Aspen.Application.Tree.FindNode(f"{henry_path}\\VAL{str(i+1)}\\HENRY\\{gas}\\{solvent}").Value = val
        
    def set_NRTL(self, comp_1, comp_2, val_list): # [A12, A21, B12, B21, C12, C21, D12, D21, E12, E21, F12, F21, T_l, T_u]
        Aspen = self.Aspen
        NRTL_path = f"\\Data\\Properties\\Parameters\\Binary Interaction\\NRTL-1\\Input"
        tpath = NRTL_path + f"\\TUNITLABEL2\\NRTL\\{comp_1}\\{comp_2}"
        
        Aspen.Application.Tree.FindNode(f"{NRTL_path}\\VAL1\\NRTL").Elements.InsertRow(0,0)
        Aspen.Application.Tree.FindNode(f"{NRTL_path}\\VAL1\\NRTL").Elements.LabelNode(0,0)[0].Value = comp_1
        Aspen.Application.Tree.FindNode(f"{NRTL_path}\\VAL1\\NRTL").Elements.LabelNode(1,0)[0].Value = comp_2
        
        # change unit to 'K'
        Aspen.Application.Tree.FindNode(tpath).Value = "K"
        
        for i, val in enumerate(val_list):
            
            valpath = NRTL_path + f"\\VAL{str(i+1)}\\NRTL\\{comp_1}\\{comp_2}"
            Aspen.Application.Tree.FindNode(valpath).Value = val

    def clear_NRTL(self):
        Aspen = self.Aspen
        NRTL_path = f"\\Data\\Properties\\Parameters\\Binary Interaction\\NRTL-1\\Input"
        Aspen.Application.Tree.FindNode(f"{NRTL_path}\\VAL1\\NRTL").RemoveAll()

class ElectrolytePropModel:
    def __init__(self, Aspen, Asp, tool):
        self.Aspen = Aspen
        self.Asp = Asp
        self.tool = tool

    # Pure properties
    def set_ion_properties(self, alias, h, g, mw, charge, itype):
        # need click_Review first
        # setup ion's H, G, charge, mw, type(for Cp), Pvap(nonvolatile) unit: kcal/mol
        Aspen = self.Aspen
        alias = alias.upper()
        
        review_path = '\\Data\\Properties\\Parameters\\Pure Components\\REVIEW-1\\Input\\VALUE'
        review_unit_path = "\\Data\\Properties\\Parameters\\Pure Components\\REVIEW-1\\Input\\UNITLABEL"
        
        # check if review has node
        self.tool.check_review_node()
        
        Aspen.Application.Tree.FindNode(f"{review_unit_path}\\DGFORM").Value = "kcal/mol"
        Aspen.Application.Tree.FindNode(f"{review_unit_path}\\DHFORM").Value = "kcal/mol"
        
        def insert_new_component(alias): # insert alias and clear the properties
            Aspen.Application.Tree.Data.Components.Specifications.Input.TYPE.Elements.InsertRow(0,0)
            Aspen.Application.Tree.Data.Components.Specifications.Input.TYPE.Elements.LabelNode(0,0)[0].Value = alias
            Aspen.Application.Tree.FindNode(f"\\Data\\Components\\Specifications\\Input\\ANAME1\\{alias}").Value = ''
        
        def set_ion_thermo_properties(alias, h, g, mw, charge, itype=None):
            
            Aspen.Application.Tree.FindNode(review_path).Elements.InsertRow(1,0)
            Aspen.Application.Tree.FindNode(review_path).Elements.LabelNode(1,0)[0].Value = alias
            Aspen.Application.Tree.FindNode(f"{review_path}\\DHAQFM\{alias}").Value = h
            Aspen.Application.Tree.FindNode(f"{review_path}\\DGAQFM\{alias}").Value = g
            Aspen.Application.Tree.FindNode(f"{review_path}\\MW\{alias}").Value = mw
            Aspen.Application.Tree.FindNode(f"{review_path}\\CHARGE\{alias}").Value = charge
            Aspen.Application.Tree.FindNode(f"{review_path}\\IONTYP\{alias}").Value = itype
            
        def set_ion_plxant(alias):
            plxant_path = '\\Data\\Properties\\Parameters\\Pure Components\\PLXANT-1\\Input\\VAL1\\PLXANT'
            Aspen.Application.Tree.FindNode(plxant_path).Elements.InsertRow(0,0)
            Aspen.Application.Tree.FindNode(plxant_path).Elements.LabelNode(0,0)[0].Value = alias
            Aspen.Application.Tree.FindNode(f"{plxant_path}\\{alias}").Value = -1E+20
            Aspen.Application.Tree.FindNode(f"\\Data\\Properties\\Parameters\\Pure Components\\PLXANT-1\\Input\\VAL8\\PLXANT\\{alias}").Value = -273.15
            Aspen.Application.Tree.FindNode(f"\\Data\\Properties\\Parameters\\Pure Components\\PLXANT-1\\Input\\VAL9\\PLXANT\\{alias}").Value = 1726.85

        insert_new_component(alias)
        set_ion_thermo_properties(alias, h, g, mw, charge, itype)
        set_ion_plxant(alias)

    # Binary properties
    def set_eNRTL(self, m, cation, anion, Cmca=-4, Ccam=8, Dmca=0, Dcam=0, Emca=0, Ecam=0, alpha=0.2):
        Aspen = self.Aspen
        def set_a_pair_of_eNRTL_parameters(m, c, a, pair_name, Pmca, Pcam):
            eNRTL_path = f"\\Data\\Properties\\Parameters\\Electrolyte Pair\\{pair_name}-1\\Input"
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\VALUE\\{pair_name}").Elements.InsertRow(0,0)
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\CID1A\\{pair_name}\\#0").Value = m
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\CID2A\\{pair_name}\\#0").Value = c
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\CID2B\\{pair_name}\\#0").Value = a
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\VALUE\\{pair_name}\\#0").Value = Pmca

            if pair_name == 'GMELCN': # alpha has no pair
                return
            
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\VALUE\\{pair_name}").Elements.InsertRow(0,0)
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\CID2A\\{pair_name}\\#0").Value = m
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\CID1A\\{pair_name}\\#0").Value = c
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\CID1B\\{pair_name}\\#0").Value = a
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\VALUE\\{pair_name}\\#0").Value = Pcam

        parameters = {
            'GMELCC': (Cmca, Ccam),
            'GMELCD': (Dmca, Dcam),
            'GMELCE': (Emca, Ecam),
            'GMELCN': (alpha, alpha)
        }
        
        # change unit to K
        Aspen.Application.Tree.FindNode("\\Data\\Properties\\Parameters\\Electrolyte Pair\\GMELCD-1\\Input\\UNITLABEL\\GMELCD").Value = 'K'
        
        for model, (param1, param2) in parameters.items():
            set_a_pair_of_eNRTL_parameters(m, cation, anion, model, param1, param2)

    def clear_eNRTL(self):
        Aspen = self.Aspen
        pair_names = ['GMELCC', 'GMELCD', 'GMELCE', 'GMELCN']
        for pair_name in pair_names:
            eNRTL_path = f"\\Data\\Properties\\Parameters\\Electrolyte Pair\\{pair_name}-1\\Input"
            Aspen.Application.Tree.FindNode(f"{eNRTL_path}\\VALUE\\{pair_name}").RemoveAll()

    # chemical reaction equilibrium (salts not yet)
    # reaction = {
    #         'R': {'AM+': 1, 'H2O': 1},
    #         'P': {amine: 1, 'H3O+': 1}
    # }
    def set_Chemistry(self, reaction): # not stable
        import pyperclip
        import warnings
        warnings.filterwarnings("ignore", category=RuntimeWarning)
        
        Asp = self.Asp
        Aspen = self.Aspen
        chem_name = 'YUE-CHE'
        Aspen.Application.Tree.FindNode(r"\Data\Properties\Specifications\Input\GCHEMISTRY").Value = chem_name
        def extract_reactants_and_products(reaction):
            reactants = reaction['R']
            products = reaction['P']

            # Prepare input for setup_one_side in the format (chemical, coefficient)
            reactant_items = [(k, v) for k, v in reactants.items()]
            product_items = [(k, v) for k, v in products.items()]
            
            return reactant_items, product_items
        
        def setup_one_side(one_side_list):
            for item in one_side_list:
                if not isinstance(item, tuple) or len(item) != 2:
                    raise ValueError("Each item in the list must be a tuple of chemical and coefficient.")

            formatted_rxn = ""
            
            # Processes each tuple in the list
            for chemical, coefficient in one_side_list:
                if formatted_rxn:  # If there is already content, add a newline
                    formatted_rxn += "\n"
                formatted_rxn += f"{chemical}\t{coefficient}"
            
            return formatted_rxn
        
        reactants, products = extract_reactants_and_products(reaction)
        
        input_item = Asp.child_window(title=f'{chem_name}', class_name="TabItem", control_type='TabItem')
        input_chem = input_item.child_window(title = 'Chemistry', auto_id="MMTabItem_1", class_name="TabItem")
        if not input_chem.exists(): # if not at input folder
            # cd chemistry
            Asp.child_window(title="Chemistry", auto_id="igRibbon_btnProperties_Chemistry", control_type="Button").click()
            # cd Chemistry_GLOBAL
            Asp.children()[3].children()[0].children()[3].children()[2].children()[3].children()[3].select().type_keys('{ENTER}')
        
        # add new equilibrium reaction
        Asp.child_window(title = 'New', auto_id="MMCmdButton_1", class_name="Button", control_type="Button").click()
        Asp.child_window(title = 'OK', auto_id="PART_BUTTON", class_name="Button", control_type="Button").click()
        
        reac_name = "AspenTech.AspenPlus.MMForms.Reactions.Chemistry.MMChemistry_Input_EditEqui_ViewModel+grdReac_GridData_ItemTemplate"
        prod_name = "AspenTech.AspenPlus.MMForms.Reactions.Chemistry.MMChemistry_Input_EditEqui_ViewModel+grdProd_GridData_ItemTemplate"
        
        window = Asp.child_window(title='txtRxnId', class_name="MMValueEdit", auto_id='MMValueEdit_2')
        edit_box = window.child_window(class_name="XamTextEditor", control_type='Edit', auto_id='PART_editControl')
        rxn_no = int(edit_box.get_value())
        
        Asp.child_window(title=reac_name, class_name="Record", found_index=0).set_focus().select()
        pyperclip.copy(setup_one_side(reactants))
        Asp.child_window(title=reac_name, class_name="Record", found_index=0).set_focus().select().type_keys('^V')
        # Asp.child_window(title = "CxvVectorPropertyData - 2 Fields - index = 0", class_name="Header", found_index=0).type_keys('{DOWN}^v')
        time.sleep(0.5)
        
        pyperclip.copy(setup_one_side(products))
        Asp.child_window(title=prod_name, class_name="Record", found_index=0).set_focus().select().type_keys('^V')
        # Asp.child_window(title = "CxvVectorPropertyData - 2 Fields - index = 0", class_name="Header", found_index=1).type_keys('{DOWN}^v')
        Asp.child_window(title = "Close", auto_id="PART_BUTTON", class_name="Button", control_type="Button").click()
        
        rxn_string = Aspen.Application.Tree.FindNode(f"\\Data\\Reactions\\Chemistry\\{chem_name}\\Input\\RXNSTRING\\#{str(rxn_no-1)}").Value
        print(rxn_string)

class SolidPropModel:
    def __init__(self, Aspen, Asp, tool):
        self.Aspen = Aspen
        self.Asp = Asp
        self.tool = tool
        

class Simulation:
    def __init__(self, Aspen, Asp, tool):
        self.Aspen = Aspen
        self.Asp = Asp
        self.tool = tool

    def blk(self):
        return self.Aspen.Application.Tree.Elements("Data").Elements("Blocks")

    def strm(self):
        return self.Aspen.Application.Tree.Elements("Data").Elements("Streams")

    def Run(self): # run simulation
        Asp = self.Asp
        tool_bar = Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", class_name='QuickAccessToolbar', control_type="ToolBar")
        run = tool_bar.child_window(title="Run", auto_id='igRibbon_btnRun', control_type="Button")
        if run.legacy_properties()['State'] != 1:
            run.click()
            time.sleep(1.6)

    def Reset(self): # reset simulation
        Asp = self.Asp
        def click_two_oks():
            window_1 = Asp.child_window(title="Reinitialize", auto_id="Window_1", class_name="Window")
            ok_1 = window_1.child_window(title="OK", auto_id="Button_1", control_type="Button")
            ok_1.click()
            time.sleep(0.1)
            window_2 = Asp.child_window(title="Aspen Plus", auto_id="Window_1", class_name="Window")
            ok_2 = window_2.child_window(title="OK", auto_id="btn0", control_type="Button")
            ok_2.click()
        
        tool_bar = Asp.child_window(auto_id="igRibbon_QuickAccessToolbar_1", class_name='QuickAccessToolbar', control_type="ToolBar")
        reset = tool_bar.child_window(title="Reset", auto_id='igRibbon_btnReset', control_type="Button")

        try:
            reset.click()
            click_two_oks()
        except:
            return

    def results_available(self): # return True or False
        Asp = self.Asp
        results_available = Asp.child_window(title="Results Available", auto_id="PART_StatusText", class_name="TextBlock").exists()
        results_available_with_errors = Asp.child_window(title="Results Available with Errors", auto_id="PART_StatusText", class_name="TextBlock").exists()
        return results_available

    def place_block(self, block_name: str, block_type: Literal["RCSTR", "RPlug", "DSTWU", "Flash2", "Mixer", "Heater", "Radfrac", "Splitter", "RYield"]):
        self.blk().Elements.Add(f'{block_name}!{block_type}')

    def place_stream(self, stream_name: str, stream_type: Literal["MATERIAL", "HEAT", ""] = "MATERIAL"):
        self.strm().Elements.Add(f'{stream_name}!{stream_type}')

    def connect_stream(self, block_name:str, stream_name:str, port_name: Literal["V(OUT)", "L(OUT)", "F(IN)"]):
        self.blk().Elements(block_name).Elements("Ports").Elements(port_name).Elements.Add(stream_name)

    def remove_blk_strm(self):
        self.blk().RemoveAll()
        self.strm().RemoveAll()

    def flash_type(self, block_name, flash_type: Literal["TP","TD","TV","TQ","PD","PV","PQ"]='TP'):
        self.blk().Elements(block_name).Elements("Input").Elements("SPEC_OPT").Value = flash_type

    def tune_flash2(self, block_name, option: Literal['TEMP', 'PRES', 'VFRAC'], value):
        """ Unit
        Temperature: C
        Pressuse: bar
        """
        self.blk().Elements(block_name).Elements("Input").Elements(option).Value = value
        
    def tune_stream(self, stream_name, temp=25, pres=1, totflow=1, 
                    option: Literal['MASS-FLOW', 'MOLE-FLOW', 'STDVOL-FLOW', 'MASS-FRAC', 'MOLE-FRAC', 'STDVOL-FRAC', 'MASS-CONC', 'MOLE-CONC']='MOLE-FRAC'):
        """ Unit
        Temperature: C
        Pressuse: bar
        Flow: kmol/hr
        """
        strm_input = self.strm().Elements(stream_name).Elements("Input")

        strm_input.Elements("BASIS").Elements("MIXED").Value = option
        strm_input.Elements("TEMP").Elements("MIXED").Value = temp
        strm_input.Elements("PRES").Elements("MIXED").Value = pres
        # strm.Elements(stream_name).Elements("Input").Elements("VFRAC").Elements("MIXED").Value = vfrac
        if option in ['MASS-FRAC', 'MOLE-FRAC', 'STDVOL-FRAC']:
            strm_input.Elements("TOTFLOW").Elements("MIXED").Value = totflow
        else:
            strm_input.Elements("TOTFLOW").Elements("MIXED").Value = ''

    def set_stream_composition(self, stream_name, comp, value):
        self.strm().Elements(stream_name).Elements("Input").Elements("FLOW").Elements("MIXED").Elements(comp).Value = value

    # read the input species' flow rate
    def in_comp(self, stream_name, comp):
        total_flow_rate = 0
        flow = self.strm().Elements(stream_name).Elements("Input")
        
        for c in self.tool.comp_list():
            c_flow_rate = flow.Elements("FLOW").Elements('MIXED').Elements(c).Value
            if c_flow_rate != None:
                total_flow_rate += c_flow_rate

        comp_flow_rate = flow.Elements("FLOW").Elements('MIXED').Elements(comp).Value
        
        basis = flow.Elements("BASIS").Elements('MIXED').Value
        ''' basis option:
        MASS-FLOW
        MOLE-FLOW
        STDVOL-FLOW
        MASS-FRAC
        MOLE-FRAC
        STDVOL-FRAC
        MASS-CONC
        MOLE-CONC
        '''
        flow_base = flow.Elements("FLOWBASE").Elements('MIXED').Value
        '''flow base option:
        MASS
        MOLE
        STDVOL
        VOLUME
        '''
        flow_rate = 1
        
        if 'FLOW' in basis:
            return comp_flow_rate
        elif 'FRAC' in basis:
            return comp_flow_rate / total_flow_rate
        
        return comp_flow_rate

    # read the output species' flow rate
    def out_comp(self, stream_name, comp, 
                    option: Literal['MASSFLOW', 'MOLEFLOW', 'MASSFRAC', 'MOLEFRAC']):
        """
        mole flow, kmol/hr
        mass flow, kg/hr
        """
        return self.strm().Elements(stream_name).Elements("Output").Elements(option).Elements('MIXED').Elements(comp).Value

    # read the output's property (temperature, pressure, mole flows, mass flows)
    def out_props(self, stream_name, option: Literal['TEM_OUT', 'PRES_OUT', 'MOLEFLMX', 'MASSFLMX']):
        """
        TEM_OUT: temperature, K
        PRES_OUT: pressure, bar
        MOLEFLMX: total mole flow, kmol/hr
        MASSFLMX: total mass flow, kg/hr
        """
        return self.strm().Elements(stream_name).Elements("Output").Elements(option).Elements('MIXED').Value