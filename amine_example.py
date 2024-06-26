#%%
import sys
import os
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, LogLocator, AutoLocator

sys.path.append(os.path.abspath('C:\\Users\\Yang\\Desktop\\program\\aspen'))

from aspen_plus_model import startup_aspen_plus, kill_aspen_plus
from aspen_plus_model import Tool, ElectrolytePropModel, FluidPropModel, Simulation

kill_aspen_plus()

aspen_filename = r'C:\Users\Yang\Desktop\112-2\amine\ELECNRTL_Rate_Based_MEA_Model_no_change.apwz'
Aspen, Asp = startup_aspen_plus(aspen_filename)
tool = Tool(Aspen, Asp)
e_mod = ElectrolytePropModel(Aspen, Asp, tool)
f_mod = FluidPropModel(Aspen, Asp, tool)
sim = Simulation(Aspen, Asp, tool)

#%%
def collect_co2(stream): # kmol/sec
    flow_rate_co2 = 0
    for c_comp in ['CO2', 'HCO3-', 'CO3-2', 'MEACOO-']:
        flow_rate_co2  += sim.out_comp(stream, c_comp, 'MOLEFLOW')
    return flow_rate_co2

def collect_amine(stream):
    flow_rate_amine = 0
    for c_comp in ['MEA', 'MEAH+', 'MEACOO-']:
        flow_rate_amine  += sim.out_comp(stream, c_comp, 'MASSFRAC')
    return flow_rate_amine

leanin_co2 = collect_co2('LEANIN')
fluegas_co2 = collect_co2('FLUEGAS')
gasout_co2 = collect_co2('GASOUT')
co2out_co2 = collect_co2('CO2OUT')
leanout_co2 = collect_co2('LEANOUT')

input_co2 = leanin_co2 + fluegas_co2
treated_co2 = co2out_co2
untreated_co2 = gasout_co2 + leanout_co2

input_mea = collect_amine('LEANIN')

print(input_mea)
print(leanin_co2 + fluegas_co2)
print(gasout_co2 + co2out_co2 + leanout_co2)
# %%
