import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

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

#Combine with total
generation_combine = pd.merge(data_combine, state_generation, on='State', how='outer')
#Fill 0 to those states don't do wind generation
generation_combine = generation_combine.fillna(0)
#Calculate total % of generation from wind (in %)
generation_combine['Wind_percent'] = (generation_combine['Wind_Total'] / generation_combine['Total']) * 100

#Show the figure
plt.figure(figsize=(12, 8))
plt.scatter(generation_combine['Wind_percent'], generation_combine['Capacity_Factor'])

slope, intercept, r_value, p_value, std_err = stats.linregress(generation_combine['Wind_percent'], generation_combine['Capacity_Factor'])
line = slope * generation_combine['Wind_percent'] + intercept
plt.plot(generation_combine['Wind_percent'], line, color='red', label=f'Line of Best Fit (R^2={r_value**2:.2f})')


plt.xlabel('Generation from Wind %')
plt.ylabel('Capacity Factor %')
plt.title('Scatter Plot for Comparing Generation from Wind and Wind Capacity Factor')
plt.legend()

#Add caption
caption1 = 'Figure 1: This figure displays the capacity factor of each U.S. state. \nThe differences in wind capacity factor among states vary significantly depending on their reliance on and percentage of generation from wind.'
plt.text(0, -0.1, caption1, transform=plt.gca().transAxes, fontsize=10, ha='left', va='center', multialignment='left')
plt.show()