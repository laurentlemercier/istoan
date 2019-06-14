import pandas as pd
import numpy as np
import tkinter as tk
import webbrowser
import time
import argparse
import os
import glob
import shutil

ticker = ''
year_current, month = time.strftime("%Y,%m").split(',')
year_last = int(year_current) - 1
user_path = os.getenv('HOMEDRIVE') + os.getenv('HOMEPATH')
archive_path = user_path + "/Google Drive/Bourse/"

class TickerWin:
	def __init__(self):
		self.win = tk.Tk()
		self.win.title("Hello")
		tk.Label(self.win,text="Ticker Bloomberg").grid(row=0)
		self.winTicker = tk.StringVar()
		tickerEntry = tk.Entry(self.win,textvariable=self.winTicker).grid(row=1)
		tk.Button(self.win,text="Submit",command=self.show_ticker).grid(row=2)
		self.win.mainloop()

	def show_ticker(self):
		print("Ticker: %s" % (self.winTicker.get()))
		self.win.quit()
	
	def get_ticker(self):
		return self.winTicker.get()

class Action():
	def __init__(self, ticker):
		self.ticker_ft = ticker
		self.ticker = ticker.split(':')[0]
		self.name = ''
		self.index = ''
		self.url_gg = ''
		self.url_ft = "https://markets.ft.com/data/equities/tearsheet/summary?s=" + ticker
		self.url_ih = ''
		self.url_wj = ''
		self.url_bb = ''
		self.url_ms_kr = ''
		self.url_ms_is = ''
		self.url_ms_bs = ''
		self.url_ms_cs = ''
		
	def set_name_index(self, name, index):
		self.name = name
		self.index = index
		print("Name: "+self.name+" - Index: "+self.index)
		
	def open_urls(self):
		region_ft = self.ticker_ft.split(':')[1]
		comp_ft = self.ticker_ft.split(':')[0]
		market_ft = ['PAR','BRU','AEX']
		market_ind = market_ft.index(region_ft)
		market_bb = ['FP','BB','NA']
		market_ms = ['fra','bel','net']
		market_ih = ['fp','bb','na']
		market_wj = ['FR/XPAR','BE/XBRU','NL/XAMS']
		
		self.url_ih = "https://www.devenir-rentier.fr/bourse_actions/actions/"+market_ih[market_ind]+"/"+comp_ft.lower()+"_"+market_ih[market_ind]+".php"
		self.url_wj = "https://quotes.wsj.com/"+market_wj[market_ind]+"/"+comp_ft+"/financials"
		self.url_gg = "https://www.google.com/search?q=pdf+registration+document+annual+report+" + str(year_last) + "+" + self.name
		self.url_bb = "https://www.bloomberg.com/quote/"+comp_ft+":"+market_bb[market_ind]
		self.url_ms_kr = "http://financials.morningstar.com/ratios/r.html?t="+comp_ft+"&region="+market_ms[market_ind]+"&culture=en-US"
		self.url_ms_is = "http://financials.morningstar.com/income-statement/is.html?t="+comp_ft+"&region="+market_ms[market_ind]+"&culture=en-US"
		self.url_ms_bs = "http://financials.morningstar.com/balance-sheet/bs.html?t="+comp_ft+"&region="+market_ms[market_ind]+"&culture=en-US"
		#self.url_ms_cs = "http://financials.morningstar.com/cash-flow/cf.html?t="+comp_ft+"&region="+market_ms[market_ind]+"&culture=en-US"
		
		chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
		print("Open: "+self.url_gg)
		webbrowser.get(chrome_path).open_new(self.url_gg)
		#webbrowser.get(chrome_path).open_new(self.url_ih)
		#webbrowser.get(chrome_path).open_new(self.url_bb)
		webbrowser.get(chrome_path).open_new(self.url_ms_kr)
		webbrowser.get(chrome_path).open_new(self.url_ms_is)
		webbrowser.get(chrome_path).open_new(self.url_ms_bs)
		#webbrowser.get(chrome_path).open_new(self.url_ms_cs)
		#webbrowser.get(chrome_path).open_new(self.url_wj)
		#webbrowser.get(chrome_path).open_new(self.url_ft)
		
	def archive_files(self):
		download_path = user_path + "/Downloads/"
		report_path = archive_path + "1-Rapports/"
		ms_path = archive_path + "2-MorningStar/"
		
		# Archive Registration document in Archive Path
		list_of_pdf = glob.glob(download_path+'*.pdf')
		src_pdf = max(list_of_pdf, key=os.path.getctime)
		print(src_pdf)
		dst_pdf = str(year_last)+"-"+self.index+"-"+self.name+"-Registration Document.pdf"
		print ("Moving: %s to %s" %(src_pdf,report_path+dst_pdf)) 
		shutil.move(src_pdf,report_path+"/"+dst_pdf)
		
		# Archive MorningStar CSVs in Archive Path
		ms_kr_file = self.ticker + " Key Ratios.csv"
		ms_is_file = self.ticker + " Income Statement.csv"
		ms_bs_file = self.ticker + " Balance Sheet.csv"
		ms_cs_file = self.ticker + " Cash Flow.csv"
		
		print("Moving MS files in: "+ms_path)
		shutil.move(download_path+ms_kr_file,ms_path+ms_kr_file)
		shutil.move(download_path+ms_is_file,ms_path+ms_is_file)
		shutil.move(download_path+ms_bs_file,ms_path+ms_bs_file)
		#shutil.move(download_path+ms_cs_file,ms_path+ms_cs_file)
		
	def ms_files_concat(self):
		ms_path = archive_path + "2-MorningStar/"
		
		ms_kr_file = ms_path + self.ticker + " Key Ratios.csv"
		ms_is_file = ms_path + self.ticker + " Income Statement.csv"
		ms_bs_file = ms_path + self.ticker + " Balance Sheet.csv"
		#ms_cs_file = ms_path + self.ticker + " Cash Flow.csv"
		
		ms_kr_pd_name = ['Critere','TTM', str(year_last), str(year_last-1),  str(year_last-2),  str(year_last-3), str(year_last-4), str(year_last-5),  str(year_last-6),  str(year_last-7),  str(year_last-8),  str(year_last-9)]
		ms_kr_pd = pd.read_csv(ms_kr_file,header=None,skiprows=3,names=ms_kr_pd_name,decimal='.')
		ms_kr_pd.replace(regex=True,inplace=True,to_replace=r',',value=r'')
		ms_is_pd = pd.read_csv(ms_is_file,header=None,skiprows=2,names=ms_kr_pd_name,decimal='.')
		ms_bs_pd_name = ['Critere', str(year_last), str(year_last-1),  str(year_last-2),  str(year_last-3), str(year_last-4), str(year_last-5),  str(year_last-6),  str(year_last-7),  str(year_last-8),  str(year_last-9)]
		ms_bs_pd = pd.read_csv(ms_bs_file,header=None,skiprows=2,names=ms_bs_pd_name,decimal='.')
		ms_bs_pd.insert(1,'TTM',np.nan)
		
		#ms_cs_pd = pd.read_csv(ms_cs_file,header=None,skiprows=2,names=ms_kr_pd_name,decimal='.')
		
		ms_pd = pd.concat([ms_kr_pd,ms_is_pd,ms_bs_pd],ignore_index=True)
		#print(ms_pd)
		ms_pd.to_csv(ms_path+self.ticker + "_MS.csv",sep=';',decimal=',',index=False)
		
		excel_path = "C:\Program Files (x86)\Microsoft Office\Office12"
		cmd_ms = 'start "'+excel_path+"\EXCEL.exe"+'" "'+ms_path+self.ticker + "_MS.csv" + '" &'
		print(cmd_ms)
		os.system(cmd_ms)
		
	def write_csv_from_ms(self):
		ms_path = archive_path + "2-MorningStar/"
		cp_path = archive_path+"3-Companies/"
		cp_file = str(year_last)+"-"+self.index+"-"+self.name+"-analysis.xlsx"
		ms_pd = pd.read_csv(ms_path+self.ticker + "_MS.csv",sep=";",decimal='.')
		
		ms_pd.replace(regex=True,inplace=True,to_replace=np.nan,value=r'')
		
		my_pd = pd.DataFrame(columns=ms_pd.columns)
		my_pd = my_pd.append(ms_pd.iloc[6],ignore_index=True) #Dividendes
		my_pd = my_pd.append(ms_pd.iloc[0],ignore_index=True) #Sales
		my_pd = my_pd.append(ms_pd.iloc[5],ignore_index=True) #EPS
		#my_pd = my_pd.append(ms_pd.iloc[168],ignore_index=True) #Equity
		my_pd = my_pd.append(ms_pd[ms_pd['Critere'].str.match("Total stockholders' equity")],ignore_index=True) #Equity
		#my_pd = my_pd.append(ms_pd.iloc[156],ignore_index=True) #Debt LT
		my_pd = my_pd.append(ms_pd[ms_pd['Critere'].str.match("Long-term debt")],ignore_index=True)
		my_pd = my_pd.append(ms_pd.iloc[10],ignore_index=True) #CFO
		my_pd = my_pd.append({'Critere':'ROE'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'ROIC'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'ROCE'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'Other ratios'},ignore_index=True)
		my_pd = my_pd.append(ms_pd.iloc[8],ignore_index=True) #Shares
		my_pd = my_pd.append(ms_pd.iloc[102],ignore_index=True) #COSG
		my_pd = my_pd.append(ms_pd.iloc[1],ignore_index=True) #Gross Margin
		my_pd = my_pd.append(ms_pd.iloc[2],ignore_index=True) #Oper Income
		my_pd = my_pd.append(ms_pd.iloc[27],ignore_index=True) #Tax rate
		my_pd = my_pd.append(ms_pd.iloc[4],ignore_index=True) #Net Income
		#my_pd = my_pd.append(ms_pd.iloc[145],ignore_index=True) #Total Assets
		my_pd = my_pd.append(ms_pd[ms_pd['Critere'].str.match("Total assets")],ignore_index=True)
		#my_pd = my_pd.append(ms_pd.iloc[133],ignore_index=True) #Current Assets
		my_pd = my_pd.append(ms_pd[ms_pd['Critere'].str.match("Total current assets")],ignore_index=True)
		#my_pd = my_pd.append(ms_pd.iloc[154],ignore_index=True) #Current Liabilites
		my_pd = my_pd.append(ms_pd[ms_pd['Critere'].str.match("Total current liabilities")],ignore_index=True)
		my_pd = my_pd.append(ms_pd.iloc[87],ignore_index=True) #Current Ratio
		#my_pd = my_pd.append(ms_pd.iloc[150],ignore_index=True) #Debt ST
		my_pd = my_pd.append(ms_pd[ms_pd['Critere'].str.match("Short-term debt")],ignore_index=True)
		my_pd = my_pd.append(ms_pd.iloc[29],ignore_index=True) #Asset Turnover
		my_pd = my_pd.append(ms_pd.iloc[30],ignore_index=True) #ROA
		my_pd = my_pd.append(ms_pd.iloc[32],ignore_index=True) #ROE
		my_pd = my_pd.append(ms_pd.iloc[33],ignore_index=True) #ROIC
		my_pd = my_pd.append({'Critere':'Calculs verification Excel'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'Gross Margin'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'Current Ratio'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'Asset Turnover'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'ROA'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'ROE'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'NOPAT'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'ROIC'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'ROCE'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'Debt/Equity'},ignore_index=True)
		my_pd = my_pd.append({'Critere':'Financials - Loan'},ignore_index=True)
		my_pd.insert(1,'Comment',pd.Series(['Dividendes','Sales','EPS','Equity','Debt LT', 'CFO','','','','','Shares','COSG','Gross margin','Oper income','Tax rate','Net income','Total assets','Current assets','Current Liabilites','Current ratio','Debt ST','Asset turnover','ROA','ROE','ROIC']))
		my_pd = my_pd.drop(['TTM'],axis=1)
		print(my_pd)
		my_pd.insert(0,'Indice',pd.Series([self.index,self.index,self.index,self.index,self.index,self.index,self.index,self.index,self.index]))
		my_pd.insert(1,'Raison Sociale',pd.Series([self.name,self.name,self.name,self.name,self.name,self.name,self.name,self.name,self.name]))
		my_pd.insert(2,'Critere TTM',pd.Series(['IH Score','Piotroski','Rule#1','Rdt','P/E','P/S','P/B','Debt/Equity','Comment']))
		my_pd.insert(3,'Valeur TTM',np.nan)
		my_pd.insert(4,'Critere 2 ans',pd.Series(['Shares','COSG','Gross Margin','Current Assets','Current Liabilites','Current Ratio','Asset Turnover','ROA','CFO/Assets[1]']))
		my_pd.insert(5,'Valeur N',np.nan)
		my_pd.insert(6,'Valeur N-1',np.nan)
		my_pd.insert(9,'V.1an',np.nan)
		my_pd.insert(10,'V.5an',np.nan)
		my_pd.insert(11,'Min.5an',np.nan)
		
		excel_path = "C:\Program Files (x86)\Microsoft Office\Office12"
		cmd_dash = 'start "'+excel_path+"\EXCEL.exe"+'" "'+cp_path+cp_file+ '" &'
		print(cmd_dash)
		
		cp_temp = "0-Action-template.xlsx"
		cmd_temp = 'start "'+excel_path+"\EXCEL.exe"+'" "'+cp_path+cp_temp+ '" &'
		print(cmd_temp)
		
		
		my_pd.to_excel(cp_path+cp_file,index=False)
		os.system(cmd_dash)
		os.system(cmd_temp)

	#def update_excel(self):
	#	cp_path = archive_path+"3-Companies/"
	#	#cp_file = str(year_last)+"-"+self.index+"-"+self.name+"-analysis.xlsx"
	#	cp_file = "0-Action-template.xlsx"
	#	cp_pd = pd.read_excel(cp_path+cp_file)
	#	print(cp_pd)
		
		
		
		
		
		

parser = argparse.ArgumentParser()
parser.add_argument("-t","--ticker",dest='ticker')
parser.add_argument("-u","--url",dest='urls',action='store_true')
parser.add_argument("-a","--arch",dest='archive',action='store_true')
parser.add_argument("-c","--concat",dest='concat',action='store_true')
parser.add_argument("-w","--write",dest='write',action='store_true')
args = parser.parse_args()

if args.ticker != None:
	ticker = args.ticker

if ticker == "":
	cTicker = TickerWin()
	ticker = cTicker.get_ticker()
	#ticker = "AC:PAR" #default value for tests

myAction = Action(ticker)
actions = pd.read_csv("actions.csv",sep=';')
actions = actions.sort_values('TickerFT',ascending=False)
actions = actions.drop_duplicates(['TickerFT'],keep='first')
action = actions.loc[actions['TickerFT'] == ticker]
print(action)
myAction.set_name_index(action.iloc[0]['Nom'],action.iloc[0]['Indice'])

if args.urls is True:
	myAction.open_urls()
if args.archive is True:
	myAction.archive_files()
if args.concat is True:
	myAction.ms_files_concat()
if args.write is True:
	myAction.write_csv_from_ms()


print("End ...")