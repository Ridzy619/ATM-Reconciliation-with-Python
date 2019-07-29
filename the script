
import pandas as pd

fep_df = pd.read_csv(r"C:\Users\user\Downloads\ATM_FEP.csv", sep = ',')
gl_df = pd.read_csv(r"C:\Users\user\Downloads\gl_009 (1).csv", sep = ',', header = 2)
fep_df.drop(fep_df.columns[[0,3,7,10,11]], axis = 1, inplace = True)
gl_df.drop(gl_df.columns[[0,2,5,10,11,12]], axis = 1, inplace = True)
#"C:\Users\user\Downloads\GDP_1.csv"

#gl_df.head()

#Filter fep_df dataframe to get rows between previous unload and new unload

datetime_prev = input("Date and time of the previous unload using 24 hour time system: ")
datetime_new = input("Date and time of the new unload using 24 hour time system: ")
date_filter = (pd.to_datetime(fep_df["datetime_req1"])>pd.to_datetime(datetime_prev)) & (pd.to_datetime(fep_df["datetime_req1"]) < pd.to_datetime(datetime_new))
date_filter.tail(25)
filtered_fep_df = fep_df[date_filter]

#filtered_fep_df.head()

#Get closing gl balance from gl_df
gl_bal = gl_df[gl_df.columns[-1]].iloc[-1]
if "(" in gl_bal:
    gl_bal = gl_bal.strip('()')
    gl_bal = "-" + gl_bal
else:
    gl_bal = gl_bal.strip('()')

gl_bal = int(gl_bal.replace(',', '')[:-3])
gl_bal

#Get sum of dispensed cash from fep_df
tot_cash_fep = filtered_fep_df['amount_cash_final1'].sum()
print(tot_cash_fep)

#Calculate the total cash dispensed based on previous load and new cash unload
prev_load = 5500000
unload_tot_1000 = 468000
unload_tot_500 = 357500
tot_cash_unload = unload_tot_1000 + unload_tot_500
tot_cash_disp = prev_load - tot_cash_unload
branch_name = "Branch name goes here"
gl_number = "GL number goes here"
cash_diff = tot_cash_fep - tot_cash_disp #Get the difference between the tot_cash_disp and tot_cash_fep


branch_name
gl_number


#Compare total cash dispensed on fep to total cash dispensed based on load and unload
double_stan_df=pd.DataFrame()
i = 0
if tot_cash_fep == tot_cash_disp:
    exp_of_diff = 0 #"Your cash is balanced"
elif tot_cash_fep > tot_cash_disp:
    double_stan_df = filtered_fep_df[filtered_fep_df.duplicated(['system_trace_audit_nr1'], keep = 'last')]
    exp_of_diff = double_stan_df['amount_cash_final1'].sum()
else:
    exp_of_diff = input("Kindly explain the reason for the shortage of %s in your cash" %cash_diff)


print(cash_diff)
print(exp_of_diff)
#Sum the amounts to get the unsuccessful transactions that need to be backed out of the tot_cash_fep


#Get transactions after cash count before EOD using datetime_new
#Then remove transactions with double STAN
trxns_before_eod_df = fep_df[pd.to_datetime(fep_df["datetime_req1"])>\
                             pd.to_datetime(filtered_fep_df["datetime_req1"].iloc[-1])]
double_stan_eod_filter = trxns_before_eod_df.duplicated(['system_trace_audit_nr1'], keep = False)
trxns_before_eod_df[~double_stan_eod_filter]

#Get the sum of withdrawals after cash count before eod
tot_cash_before_eod = trxns_before_eod_df['amount_cash_final1'].sum()
print(tot_cash_before_eod)
#trxns_before_eod_df[~double_stan_eod_filter]

#Get on fep not on gl transactions
rrn_gl_col = gl_df['NARRATIVE'].str[-12:]
rrn_gl_col=pd.DataFrame(rrn_gl_col)
on_fep_filter = ~filtered_fep_df.retrieval_reference_nr1.isin(rrn_gl_col.NARRATIVE) &\
                (filtered_fep_df['amount_cash_final1']!= 0)

onfep_notongl_df = filtered_fep_df[on_fep_filter].head()
tot_onfep_notongl = onfep_notongl_df['amount_cash_final1'].sum()

df = pd.read_excel (r".\for recon.xlsx", header = 8, index_col = False, usecols = 4, keep_default_na =False)

# # dict_ = 
# # fillna(value = "")
# df.to_dict()
# certificate_df1 = pd.DataFrame.from_dict(certificate_df1)

# certificate_df1

certificate_df2
df['Unnamed: 1'] = ''
df['Unnamed: 2'] = ''
df['Unnamed: 3']= ''
df['Unnamed: 4'] = ''
# df.to_dict()

certificate_df1= pd.DataFrame.from_dict(
    {'Unnamed:1':{0:'ATM CASH COUNT',1:'BRANCH:',2:'LEDGER CODE:',3:'DATE:',4:'TIME:'},
    'Unnamed:2': {0:'',1:'',2:'',3:'',4:''}}
)


certificate_df2= pd.DataFrame.from_dict(
    {
    'Denomination (N)': {1: 1000,2: 500,3: 'Retract Bin: N1,000',4: 'Retract Bin: N500',
           5: 'Total Physical Cash',6: 'Closing balance',7: 'Difference- (Shortage)/ Overage*',8: '',
           9: 'Explaination on the differences',10: '',11: 'CASH LOAD',12: 'CASH UNLOAD',13: '',14: '',
           15: 'BSH',16: '',17: 'Back Up',18: '',19: 'Internal Control Officer'},
     'Bundle': {1: '',2: '',3: '',4: '',5: '',6: '',7: '',8: '',9: '',10: '',11: '',12: '',13: '',
                    14: '',15: '',16: '',17: '',18: '',19: ''},
     'Packet': {1: '',2: '',3: '',4: '',5: '',6: '',7: '',8: '',9: '',10: '',11: '',12: '',13: '',
                    14: '',15: '',16: '',17: '',18: '',19: ''},
     'Pieces': {1: '',2: '',3: '',4: '',5: '',6: '',7: '',8: '',9: '',10: '',11: '',12: '',13: '',
                    14: '',15: '',16: '',17: '',18: '',19: ''},
     'Amount (N)': {1: '',2: '',3: '',4: '',5: '',6: '',7: '',8: '',9: '',10: '',11: '',12: '',13: '',
                    14: '',15: '',16: '',17: '',18: '',19: ''}
    }
)


#Prepared by:
prepared_by = "Prepared by"
reviewed_by = "Reviewed by"
internal_control = "Internal Control"
date_last_load = pd.to_datetime(datetime_prev).date()
time_last_load = pd.to_datetime(datetime_prev).time()#"Time last load goes here"
date_new_load = pd.to_datetime(datetime_new).date()#"Date new load goes here"
time_new_load = pd.to_datetime(datetime_new).time()#"Time new load goes here"
branch_name
gl_number

#RECON page variables
gl_bal
tot_cash_fep
tot_recon_page = gl_bal + tot_cash_unload
net_cash_load = tot_recon_page - prev_load
tot_cash_before_eod
tot_onfep_notongl
tot_ongl_notonfep = 0
tot_onus_disperr = 0
tot_ofuss_disperr = 0
tot_disperr = tot_onus_disperr + tot_ofuss_disperr
adj_gl_bal = net_cash_load + tot_cash_before_eod + tot_ongl_notonfep - tot_onfep_notongl
tot_cash_unload
tot_cumm_diff = tot_cash_unload - adj_gl_bal
prev_cumm_diff = 0 #To be inputed
diff_today = tot_cumm_diff - prev_cumm_diff
diff_today


#Suspected Dispense Error sheet 
onus_disperr_df = pd.DataFrame() #These records will be fetched from the portal that confirms dispense error
offus_disperr_df = pd.DataFrame() 
#tot_onus_disperr = onus_disperr_df['amount_cash_final1']
#tot_offus_disperr = offus_disperr_df['amount_cash_final1']


#NotOnFepOnGl

#ongl_notonfep_df = df
#tot_ongl_notonfep = ongl_notonfep_df['amount_cash_final1']



pd.DataFrame(certificate_df1)

#Certificate Sheet
new_load = 7500000
prev_load = 5500000
unload_tot_1000 = 468000
unload_tot_500 = 357500
tot_cash_unload = unload_tot_1000 + unload_tot_500
tot_cash_disp = prev_load - tot_cash_unload

certificate_df2['Amount (N)'].loc[1] = unload_tot_1000
certificate_df2['Amount (N)'].loc[2] = unload_tot_500
certificate_df2['Bundle'].loc[11] = new_load
certificate_df2['Bundle'].loc[12] = tot_cash_unload
certificate_df2['Amount (N)'].loc[6] = -gl_bal
certificate_df2['Amount (N)'].loc[7] = tot_cash_unload - gl_bal
certificate_df2['Amount (N)'].loc[5] = tot_cash_unload
certificate_df2['Bundle'].loc[17] = prepared_by
certificate_df2['Bundle'].loc[15] = reviewed_by
certificate_df2['Bundle'].loc[19] = 'To be determined'

certificate_df1['Unnamed:2'].loc[1] = branch_name
certificate_df1['Unnamed:2'].loc[3] = date_new_load
certificate_df1['Unnamed:2'].loc[2] = gl_number
certificate_df1['Unnamed:2'].loc[4] = time_new_load


certificate_df1





atm_man_bal_df1=pd.DataFrame(
    {'Date':{0:'',1:'',2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:''},'S/N':{0:'A',1:'i',
2:'ii',3:'iii',4:'iv',5:'v',6:'vi',7:'B',8:'C',9:'D'},'DETAILS':{0:'Total cash loaded in ATM',1:'Reject Canister',
2:'Reject Canister',3:'Retract Canister',4:'Retract Canister',5:'ATM Unpaid Cash Canister',6:'ATM Unpaid Cash Canister',
7:'Total cash found in ATM(sum i to vi)',8:'Total cash dispensed for the period [A minus B](to be crosschecked with FEP)',
9:'Fresh cash reloaded today'},'Unnamed:3':{0:'',1:'N1,000',2:'N500',3:'N1,000',4:'N500',5:'N1,000',6:'N500',7:'',
8:'',9:''},'PCS':{0:'',1:'',2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:''},'AMOUNT(=N=)':{0:'',1:'',2:'',3:'',4:'',
5:'',6:'',7:'',8:'',9:''},'AMOUNT(=N=).1':{0:'',1:'',2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:''},
'MAKER Signature':{0:'',1:'',2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:''},'BSH Signature':{0:'',1:'',2:'',3:'',4:'',
5:'',6:'',7:'',8:'',9:''},'TIME OFF':{0:'',1:'',2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:''},'TIME ON':{0:'',1:'',
2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:''}}
)

atm_man_bal_df2=pd.DataFrame(
    {'Unnamed:0':{0:'TOTAL Dispensed (FEP)',1:'',
2:'Difference beween Total cash dispensed and FEP Cash dispensed',3:'',4:'Explanation of difference'},
'Unnamed:3':{0:'',1:'',2:'',3:'',4:''}}
)


#import datetime
amount_col = [prev_load,'','','','',unload_tot_500, unload_tot_1000,tot_cash_unload,tot_cash_disp,new_load]
#amount_col2 = [prev_load,'','','','',unload_tot_500, unload_tot_1000,'','','']
atm_man_bal_df1['Date'].iloc[0]= pd.to_datetime(datetime_prev).date()
atm_man_bal_df1['TIME OFF'].iloc[0]= pd.to_datetime(datetime_prev ).time()
atm_man_bal_df1['TIME ON'].iloc[0]= (pd.to_datetime(datetime_prev ) + pd.Timedelta(minutes = 1)).time()
atm_man_bal_df1['AMOUNT (=N=)'] = amount_col
atm_man_bal_df1
# print(datetime_prev)
# pd.to_datetime('2/4/1999 00:02') - pd.to_datetime(datetime_prev)

tot_cash_unload
prev_load

atm_man_bal_df2['Unnamed:3'].iloc[0] = tot_cash_fep
atm_man_bal_df2['Unnamed:3'].iloc[2] = cash_diff
atm_man_bal_df2['Unnamed:3'].iloc[4] = exp_of_diff
atm_man_bal_df1

# df = pd.read_excel (r".\for recon.xlsx", header = 14, index_col = False, sheet_name = 'Sheet2',\
#                     nrows = 6,usecols = [0,3], keep_default_na =False)

onus_disperr_df = pd.DataFrame()
offus_disperr_df = pd.DataFrame()

#Prepared by:
prepared_by = "Prepared by"
reviewed_by = "Reviewed by"
internal_control = "Internal Control"
date_last_load = pd.to_datetime(datetime_prev).date()
time_last_load = pd.to_datetime(datetime_prev).time()#"Time last load goes here"
date_new_load = pd.to_datetime(datetime_new).date()#"Date new load goes here"
time_new_load = pd.to_datetime(datetime_new).time()#"Time new load goes here"
date_of_prep = (pd.to_datetime(datetime_new) + pd.Timedelta(days = 1)).date()
atm_number = "ATM Number goes here"
branch_name
gl_number

gl_bal
tot_cash_fep
tot_recon_page = gl_bal + tot_cash_unload
net_cash_load = tot_recon_page - prev_load
tot_cash_before_eod
tot_onfep_notongl
tot_ongl_notonfep = 0
tot_onus_disperr = 0
tot_ofus_disperr = 0
tot_disperr = tot_onus_disperr + tot_ofus_disperr
adj_gl_bal = net_cash_load + tot_cash_before_eod + tot_ongl_notonfep - tot_onfep_notongl
tot_cash_unload
tot_cumm_diff = tot_cash_unload - adj_gl_bal
prev_cumm_diff = 0 #To be inputed
diff_today = tot_cumm_diff - prev_cumm_diff

onus_disperr_df = pd.DataFrame() #These records will be fetched from the portal that confirms dispense error
offus_disperr_df = pd.DataFrame()

recon_df = pd.DataFrame({0:{0:'',1:'',2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:'',10:'',11:'',12:'',13:'',14:'',15:'',16:'',17:'',
18:'',19:'',20:'',21:'',22:'',23:'',24:'',25:'',26:'',27:'',28:'',29:'',30:'',31:'',32:'',33:'',34:'',35:'',
36:'',37:'',38:'',39:'',40:'',41:'',42:'',43:'',44:'',45:'',46:'',47:''},1:{0:'BRANCH:',1:'LEDGER CODE:',
2:'PROOF AS AT:',3:'DATE OF PREPARATION',4:'ATM NUMBER',5:'TIME',6:'Closing Balance',7:'',8:'CASH LOAD',9:'',10:'',
11:'',12:'',13:'',14:'',15:'Cash Unload',16:'',17:'Total',18:'Net Cash Load',19:'',
20:'Total Cash Withdrawn after Cash Count before EOD',21:'',22:'ON FEP NOT ON GL',23:'',24:'ON GL NOT ON FEP',
25:'',26:'TOTAL SUSPECTED DISPENSE ERRORS',27:'',28:'ON US DISPENSE ERROR',29:'',30:'OFF US DISPENSE ERROR',
31:'',32:'Adjusted GL Balance',33:'',34:'Physical Cash Available in ATM',35:'',36:'',37:'',38:'',
39:'TOTAL CUMMULATIVE DIFFERENCE',40:'',41:"PREVIOUS DAY'S CUMMULATIVE DIFFERENCE",42:'',43:'DIFFERENCE TODAY',
44:'',45:'PREPARED BY:',46:'',47:'REVIEWED BY:'}})
recon_df = pd.DataFrame(recon_df)


col2 = {0:branch_name,1:gl_number,2:date_new_load,3:date_of_prep,4:atm_number,5:time_new_load,6:-gl_bal,7:'',
        8:new_load,9:'',10:'',11:'',12:'',13:'',14:'',15:tot_cash_unload,16:'',17:tot_recon_page,
        18:net_cash_load,19:'',20:tot_cash_before_eod,21:'',22:tot_onfep_notongl,23:'',24:tot_ongl_notonfep,
        25:'',26:tot_disperr,27:'',28:tot_onus_disperr,29:'',30:tot_ofus_disperr,31:'',32:adj_gl_bal,33:'',
        34:tot_cash_unload,35:'',36:'',37:'',38:'',39:tot_cumm_diff,40:'',41:prev_cumm_diff,42:'',43:diff_today,
        44:'',45:prepared_by,46:'',47:reviewed_by}

col1 = {0:'',1:'',2:'',3:'',4:'',5:'',6:date_new_load,7:'',8:date_new_load,9:'',10:'',11:'',12:'',13:'',14:'',
        15:date_new_load,16:'',17:date_new_load,18:date_new_load,19:'',20:date_new_load,21:'',22:date_new_load,
        23:'',24:date_new_load,25:'',26:date_new_load,27:'',28:'',29:'',30:'',31:'',32:date_new_load,33:'',
        34:date_new_load,35:'',36:'',37:'',38:'',39:date_new_load,40:'',41:date_of_prep,42:'',43:date_new_load,44:'',45:'',46:'',47:''}

# df = pd.read_excel (r'.\for recon.xlsx', header = None, index_col = False, sheet_name = 'Sheet3',\
#                     keep_default_na = False, usecols = [0,1])

recon_df[2] = col2.values()
recon_df

with pd.ExcelWriter('ATM Reconciliation.xlsx') as writer:
    certificate_df1.to_excel(writer, sheet_name = 'Certificate', header = False, startrow = 3, index = False)
    certificate_df2.to_excel(writer, sheet_name = 'Certificate', header = True, startrow = 9,\
                             index = False)
    double_stan_df.to_excel(writer, sheet_name = 'ATM MAN. BAL', index = False, startrow = 22)
    atm_man_bal_df1.to_excel(writer, sheet_name = 'ATM MAN. BAL', index = False)
    atm_man_bal_df2.to_excel(writer, sheet_name = 'ATM MAN. BAL', index = False, startrow = 14, header = False)
    filtered_fep_df.to_excel(writer, sheet_name = 'MANUAL BAL TRXNS', index = False)
    recon_df.to_excel(writer, sheet_name = 'ATM Recon Page', index = False, header = False)
    fep_df.to_excel(writer, sheet_name = 'FEP', index = False)
    gl_df.to_excel(writer, sheet_name = 'GL', index = False)
    trxns_before_eod_df.to_excel(writer, sheet_name = 'TrxnsAfterCashCountB4EOD', index = False)
    onus_disperr_df.to_excel(writer, sheet_name = 'Suspected Disp Err', index = False)
    offus_disperr_df.to_excel(writer, sheet_name = 'Suspected Disp Err', index = False, startrow = 15)
    onfep_notongl_df.to_excel(writer, sheet_name = 'OnFepNotOnGL', index = False)
    
    


workbook_obj = writer.book
worksheet_obj = writer.sheets['Certificate']
format_dict = {'font_size': '16','font_color': 'red','bold':True}
format_obj = workbook_obj.add_format(format_dict)


# try:
#     print('hi')
# except:
#     print("An error has occured")
# else:
#     print("You're welcome")
certificate_df2

wr = pd.ExcelWriter("Test.xlsx")
recon_df.style
