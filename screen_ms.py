import pandas as pd
import numpy as np
import os, time

index = 'CAC40'
criteria = 'Dividends' # Choose among Dividends, Revenue, etc... 

year_current, month = time.strftime("%Y,%m").split(',')
year_last = int(year_current) - 1
user_path = os.getenv('HOMEDRIVE') + os.getenv('HOMEPATH')
archive_path = user_path + "/Google Drive/Bourse/"
ms_path = archive_path + "2-MorningStar/"
sc_path = archive_path + "4-Screens/"

actions = pd.read_csv("actions.csv",sep=';')
actions = actions.loc[actions['Indice'] == index]

#print (actions)

columns_out = ['Index','TickerFT','Name','Sector','Criteria','Var 1-yr (%)','Min Var 10-yr (%)', str(year_last), str(year_last-1),  str(year_last-2),  str(year_last-3), str(year_last-4), str(year_last-5),  str(year_last-6),  str(year_last-7),  str(year_last-8),  str(year_last-9)]
columns_ms = ['Criteria','TTM', str(year_last), str(year_last-1),  str(year_last-2),  str(year_last-3), str(year_last-4), str(year_last-5),  str(year_last-6),  str(year_last-7),  str(year_last-8),  str(year_last-9)]

df_out = pd.DataFrame(columns=columns_out)



for ticker_ft in actions['TickerFT']:
	ticker = ticker_ft.split(':')[0]
	action = actions.loc[actions['TickerFT'] == ticker_ft]
	action_name = actions.loc[actions['TickerFT'] == ticker_ft,'Nom'].iloc[0]
	action_sector = actions.loc[actions['TickerFT'] == ticker_ft,'Sector B.'].iloc[0]
	print("Extract: "+action_name+" ...")
	ms_pd = pd.read_csv(ms_path+ticker+ " Key Ratios.csv",header=None,skiprows=3,names=columns_ms,decimal='.')
	ms_pd.replace(regex=True,inplace=True,to_replace=r',',value=r'')
	ms_pd.replace(regex=True,inplace=True,to_replace=np.nan,value=r'')
	for i in range(0,ms_pd['Criteria'].count()-2): #Convert MS data in figures format
		for j in range(0,10):
			if ms_pd.iloc[i][str(year_last-j)] != '' and ms_pd.iloc[i][str(year_last-j)].find("-") == -1 :
				ms_pd.iloc[i][str(year_last-j)]= float(ms_pd.iloc[i][str(year_last-j)])
			elif ms_pd.iloc[i][str(year_last-j)] == '':
				ms_pd.iloc[i][str(year_last-j)] = "0.0"
				ms_pd.iloc[i][str(year_last-j)]= float(ms_pd.iloc[i][str(year_last-j)])

	ms_pd = ms_pd[ms_pd['Criteria'].str.contains(criteria)].head(1).reset_index(drop=True)
	ms_pd = ms_pd.drop(['TTM'],axis=1)
	ms_pd.insert(0,'Index',pd.Series([index]))
	ms_pd.insert(1,'TickerFT',pd.Series([ticker_ft]))
	ms_pd.insert(2,'Name',pd.Series([action_name]))
	ms_pd.insert(3,'Sector',pd.Series([action_sector]))
	#print(ms_pd.loc[0,str(year_last)])
	if ms_pd.loc[0,str(year_last-1)] != 0.0:
		var_1yr = (ms_pd.loc[0,str(year_last)] - ms_pd.loc[0,str(year_last-1)]) / ms_pd.loc[0,str(year_last-1)]
	else:
		var_1yr = 0.0
	min_var_10yr = 1000.0
	for i in range(0,9):
		if ms_pd.loc[0,str(year_last -i-1)] != 0.0:
			print(ms_pd.loc[0,str(year_last-i-1)])
			var_yoy = (ms_pd.loc[0,str(year_last - i)]/ms_pd.loc[0,str(year_last-i-1)]) - 1
			min_var_10yr = min(min_var_10yr,var_yoy)

	ms_pd.insert(5,'Var 1-yr (%)',pd.Series([var_1yr]))
	ms_pd.insert(6,'Min Var 10-yr (%)',pd.Series([min_var_10yr]))
	print(ms_pd)
	df_out = df_out.append(ms_pd,ignore_index=True)

print(df_out)

xls_file = str(year_last)+"-Screener_"+criteria+".xlsx"
df_out.to_excel(sc_path+xls_file,index=False)

excel_path = "C:\Program Files (x86)\Microsoft Office\Office12"
cmd_dash = 'start "'+excel_path+"\EXCEL.exe"+'" "'+sc_path+xls_file+ '" &'
print(cmd_dash)
os.system(cmd_dash)


