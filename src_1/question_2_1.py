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
wind_2019 = generation[generation['State Historical Tables for 2022\nReleased: September 2023\nNext Update: October 2024'] == 2019]
#Get 2019 total wind generation
wind_2019_total = wind_2019[(wind_2019['Unnamed: 3'] == 'Wind') & (generation['Unnamed: 2'] == 'Total Electric Power Industry')]
#Get 2019 total wind generation by state
state_wind_generation = wind_2019_total.groupby('Unnamed: 1')['Unnamed: 4'].sum().reset_index()
#Get rid of US-Total
state_wind_generation = state_wind_generation.drop(35).reset_index(drop=True)
state_wind_generation.rename(columns={'Unnamed: 1': 'State'}, inplace=True)
state_wind_generation.rename(columns={'Unnamed: 4': 'Wind_Total'}, inplace=True)

#Combine them in one sheet
data_combine = pd.merge(state_nameplate, state_wind_generation, on='State')
#Calculate capacity factor. (remember MW -> MWh)
data_combine['Capacity_Factor'] = data_combine['Wind_Total'] / (data_combine['Capacity'] * 24 * 365) * 100

#Sort
sorted = data_combine.sort_values(by='Capacity_Factor', ascending=False)

#Show the figure
plt.figure(figsize=(12, 8))
plt.bar(sorted['State'], sorted['Capacity_Factor'])
plt.xlabel('States')
plt.ylabel('Capacity Factor %)')
plt.title('State-level Capacity Factor (High to low in %)')
#Add caption
caption1 = 'Figure 1: This figure displays the capacity factor of each U.S. state. The capacity factor represents the efficiency of wind power generation.'
plt.text(0.5, -0.1, caption1, transform=plt.gca().transAxes, fontsize=10, ha='center', va='center', multialignment='center')
plt.show()