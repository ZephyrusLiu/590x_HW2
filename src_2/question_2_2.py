import pandas as pd
import matplotlib.pyplot as plt

wind_capacity = pd.read_excel("3_2_Wind_Y2019.xlsx")
generation = pd.read_excel("annual_generation_state.xls")

#Get operating status facilities
OP = wind_capacity[wind_capacity['Unnamed: 7'] == 'OP']
#Get nameplate capacity of each state
state_nameplate = OP.groupby('Unnamed: 4')['Unnamed: 12'].sum().reset_index()
state_nameplate.rename(columns={'Unnamed: 4': 'State'}, inplace=True)
state_nameplate.rename(columns={'Unnamed: 12': 'Capacity'}, inplace=True)

#Get 2019 generation
generation_2019 = generation[generation['State Historical Tables for 2022\nReleased: September 2023\nNext Update: October 2024'] == 2019]
#Get 2019 total generation
all_2019_total = generation_2019[(generation_2019['Unnamed: 3'] == 'Total') & (generation['Unnamed: 2'] == 'Total Electric Power Industry')]
#Get 2019 total wind generation
wind_2019_total = generation_2019[(generation_2019['Unnamed: 3'] == 'Wind') & (generation['Unnamed: 2'] == 'Total Electric Power Industry')]

# Get 2019 total generation by state
state_generation = all_2019_total.groupby('Unnamed: 1')['Unnamed: 4'].sum().reset_index()
#Get rid of US-Total
state_generation = state_generation.drop(44).reset_index(drop=True)
state_generation.rename(columns={'Unnamed: 1': 'State'}, inplace=True)
state_generation.rename(columns={'Unnamed: 4': 'Total'}, inplace=True)

# Get 2019 total wind generation by state
state_wind_generation = wind_2019_total.groupby('Unnamed: 1')['Unnamed: 4'].sum().reset_index()
#Get rid of US-Total
state_wind_generation = state_wind_generation.drop(35).reset_index(drop=True)
state_wind_generation.rename(columns={'Unnamed: 1': 'State'}, inplace=True)
state_wind_generation.rename(columns={'Unnamed: 4': 'Wind_Total'}, inplace=True)

#Combine wind and nameplate capacity
data_combine = pd.merge(state_nameplate, state_wind_generation, on='State')
#Calculate capacity factor. (remember MW -> MWh)
data_combine['Capacity_Factor'] = data_combine['Wind_Total'] / (data_combine['Capacity'] * 24 * 365) * 100

#Combine wind and total
generation_combine = pd.merge(state_wind_generation, state_generation, on='State', how='outer')
#Fill 0 to those states don't do wind generation
generation_combine = generation_combine.fillna(0)
#Calculate total % of generation from wind (in %)


# #Sort
# sorted = data_combine.sort_values(by='Capacity_Factor', ascending=False)


print(generation_combine)
