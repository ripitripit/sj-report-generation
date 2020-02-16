import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime
import webbrowser
import os
from tkinter import *
plt.style.use('ggplot')

#convert excel files to csv
def SJHD_toCSV():
    pd.read_excel('SJHD.xlsx', 'Page 1', index_col=None).to_csv('SJHD.csv', encoding='utf-8')
def SJHDCallStats_toCSV():
    pd.read_excel('SJ Help Desk - Previous Week.xls', 'Sheet0', index_col=None).to_csv('SJ Help Desk - Previous Week.csv', encoding='utf-8')
def SJM_toCSV():
    pd.read_excel('SJM.xlsx', 'Page 1', index_col=None).to_csv('SJM.csv', encoding='utf-8')
def SJMCallStats_toCSV():
    pd.read_excel('SJ Milli - Previous Week.xls', 'Sheet0', index_col=None).to_csv('SJ Milli - Previous Week.csv', encoding='utf-8')

def convert_toCSV():
    SJHD_toCSV()
    SJHDCallStats_toCSV()
    SJM_toCSV()
    SJMCallStats_toCSV()
    if os.path.isfile('SJHD.csv'):
        print("The file, SJHD.csv, generated successfully")
    else:
        print("The file, SJHD.csv, DID NOT generate successfully")
    if os.path.isfile('SJ Help Desk - Previous Week.csv'):
        print("The file, SJ Help Desk - Previous Week.csv, generated successfully")
    else:
        print("The file, SJ Help Desk - Previous Week.csv, DID NOT generate successfully")
    if os.path.isfile('SJM.csv'):
        print("The file, SJM.csv, generated successfully")
    else:
        print("The file, SJM.csv, DID NOT generate successfully")
    if os.path.isfile('SJ Milli - Previous Week.csv'):
        print("The file, SJ Milli - Previous Week.csv, generated successfully")
    else:
        print("The file, SJ Milli - Previous Week.csv, DID NOT generate successfully")

#find open/res/esc/created tickets
##def find(fileName):
##    df = pd.read_csv(fileName, delimiter=',', usecols=['Assignment group', 'Resolved'])
##    Helpdesk = 'IS 4th Source Helpdesk'
##    Esc_ct = df[(df['Assignment group'] != Helpdesk)].count()
##    Res_ct = df[(df['Assignment group'] == Helpdesk)].count()
##    opened = df[df['Resolved'].isnull()]
##    print('Opened: ', opened)
##    print('Escalated: ', Esc_ct)
##    print('Resolved', Res_ct)
##find('SJM.csv')

#find total calls in call statistics report
def find_total_calls(df):
    totalRows = (len(df.index - 30))
    return (totalRows - 29)

#clean list
def clean_list(df, totalCalls):
    df_toList = df.iloc[-5].tolist()
    del df_toList[0]
    clean1_list = [x for x in df_toList if str(x) != 'nan']
    for x in range(len(clean1_list)):
        for char in "Avg: ": clean1_list[x] = clean1_list[x].replace(char,"")
    clean2_list = [x for x in clean1_list if len(x) < 5 and len(x) > 0]
    clean3_list = [x for x in clean2_list if str(x) != '  ']
    clean3_list.append(totalCalls)
    return list(map(int, clean3_list))

#clean list for averages
def clean_list_averages(df):
    df_list = df.iloc[-5].tolist()
    del df_list[0]
    clean1_list = [x for x in df_list if str(x) != 'nan']
    for x in range(len(clean1_list)):
        for char in "Avg: ": clean1_list[x] = clean1_list[x].replace(char,"")
    clean2_list = [x for x in clean1_list if len(x) > 5]
    clean2_list = [x for x in clean2_list if str(x) != '  ']
    return clean2_list

def EscRes(fileName, title, imageName):
    df = pd.read_csv(fileName, delimiter=',', usecols=['Assignment group'])
    Helpdesk = 'IS 4th Source Helpdesk'
    Esc_ct = df[(df['Assignment group'] != Helpdesk)].count()
    Res_ct = df[(df['Assignment group'] == Helpdesk)].count()
    plt.bar('Esc', Esc_ct, label='Escalated', color='steelblue')
    plt.bar('Res', Res_ct, label='Resolved', color='darkorange')
    plt.title(title)
    plt.legend()
    plt.savefig("File_{}.png".format(imageName))
    plt.show()

def DailyTicketCount(fileName, imageName):
    df = pd.read_csv(fileName, delimiter=',')
    #convert dataframe to date & count frequency of days
    df = pd.to_datetime(df['Opened'])
    df1 = df.groupby(df.dt.date).count()
    print(df1)
    df2 = df1.plot(kind="bar", figsize=(14,6), fontsize=9)
    df2.set_alpha(0.8)
    df2.set_title("", fontsize=9)
    df2.set_xlabel("\nSunday - Saturday", fontsize=9)
    #df.set_xticklabels(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], rotation=45)
    for p in df2.patches:
        df2.annotate ("%.0f" % p.get_height(), (p.get_x() + p.get_width() / 2., p.get_height()), ha='center', va='center', xytext=(0, 10), textcoords='offset points')
    plt.subplots_adjust(left=None, right=None, bottom=.2, wspace=0, hspace=0)
    plt.savefig("File_{}.png".format(imageName), bbox_inches='tight')
    plt.show()

def AsgmtGroup(fileName, imageName):
    df = pd.read_csv(fileName)
    ax = df['Assignment group'].value_counts().plot(kind='barh', figsize=(14,6),color="Orange", fontsize=10);
    ax.set_alpha(0.8)
    ax.set_title("", fontsize=10)
    ax.set_xlabel(" \nTickets", fontsize=10);
    #ax.set_xticks([0, 5, 10, 15, 20, 25, 30])
    #find the plt.patches values and append to list
    totals = []
    for i in ax.patches:
        totals.append(i.get_width())
    #set individual bar lables using totals list
    total = sum(totals)
    #set individual bar labels using totals list
    for i in ax.patches:
        # get_width pulls left or right; get_y pushes up or down
        ax.text(i.get_width()+.3, i.get_y()+.38, \
                str(round((i.get_width()/total)*100, 2))+'%', fontsize=10,color='dimgrey')
    #invert for largest on top 
    ax.invert_yaxis()
    #configure subplot parameters
    plt.subplots_adjust(left=None, bottom=.25, right=None, top=None, wspace=None, hspace=None)
    plt.savefig("File_{}.png".format(imageName), bbox_inches='tight')
    #show graph
    plt.show()

def CallStat1(fileName, graphTitle, imageName):
    df = pd.read_csv(fileName, delimiter=',')
    #find total calls
    totalCalls = find_total_calls(df)
    #turn dataframe to list & clean list
    cleaned_list = clean_list(df, totalCalls)
    #configure bar chart
    newDF = pd.DataFrame(list(cleaned_list))
    xlabels = (['Calls Placed on Hold', 'Transfers', 'Voicemails', 'Abandoned Calls', 'Total Calls']) 
    ax = (newDF).plot(kind='bar', color='blueviolet', figsize=(10,5), fontsize=9)
    ax.set_title(graphTitle, fontsize=10)
    ax.set_xticklabels(xlabels, fontsize=10, rotation=45)
    ax.get_legend().remove()
    rects = ax.patches
    #configure date formatted labels
    for rect, label in zip(rects, cleaned_list):
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width() / 2, height + .5, label, ha='center', va='bottom')
    plt.yticks([])
    plt.subplots_adjust(left=.01, bottom=.4, right=.99, top=None, wspace=None, hspace=None)
    plt.savefig("File_{}.png".format(imageName))
    plt.show()

def CallStat2(Iteration, fileName, graphTitle, CSxlabels, imageName):
    df = pd.read_csv(fileName, delimiter=',')
    #find total calls
    totalCalls = find_total_calls(df)
    #turn dataframe to list & clean list
    cleaned_list = clean_list_averages(df)
    #slice strings from list, separate slices to new lists, & convert values to int
    hrs_list = [x[0:2] for x in cleaned_list]
    mins_list = [x[2:4] for x in cleaned_list]
    secs_list = [x[4:6] for x in cleaned_list]
    hrs_list = list(map(int, hrs_list))
    mins_list = list(map(int, mins_list))
    secs_list = list(map(int, secs_list))
    #calculate total times and place values in to new lists
    hrsToSecs_list = []
    minsToSecs_list = []
    totalTimeAsSecs_list = []
    times_list = []
    for x in range(len(hrs_list)):
        hrsToSecs_list.append(hrs_list[x]*3600)
    for x in range(len(mins_list)):
        minsToSecs_list.append(mins_list[x]*60)
    for x in range(Iteration):
        totalTimeAsSecs_list.append(hrsToSecs_list[x] + minsToSecs_list[x] + secs_list[x])
    for x in range(Iteration):
        times_list.append(str(datetime.timedelta(seconds=totalTimeAsSecs_list[x])))
    #configure bar chart
    newDF = pd.DataFrame(list(totalTimeAsSecs_list))
    xlabels = CSxlabels
    ax = (newDF).plot(kind='bar', color='lightseagreen', figsize=(10,5), fontsize=9)
    ax.set_title(graphTitle, fontsize=10)
    ax.set_xticklabels(xlabels, fontsize=10, rotation=45)
    ax.get_legend().remove()
    rects = ax.patches
    #configure date formatted labels
    for rect, label in zip(rects, times_list):
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width() / 2, height + 5, label, ha='center', va='bottom')
    plt.yticks([])
    plt.subplots_adjust(left=.01, bottom=.4, right=.99, top=None, wspace=None, hspace=None)
    plt.savefig("File_{}.png".format(imageName))
    plt.show()

def SJHD_TickReports():
    df = pd.read_csv('SJHD.csv',delimiter=',',usecols=['Number', 'Short description', 'Assignment group'])
    #identify esc/res tickets & place in to dataframes
    Helpdesk = 'IS 4th Source Helpdesk'
    Esc = df[(df['Assignment group'] != Helpdesk)]
    Res = df[(df['Assignment group'] == Helpdesk)]
    #convert esc/res dataframes to HTML
    Esc.to_html('SJHD_Escalated.html')
    Res.to_html('SJHD_Resolved.html')
    #open newly created html tables
    webbrowser.open_new_tab('SJHD_Escalated.html')
    webbrowser.open_new_tab('SJHD_Resolved.html')

def SJM_TickReports():
    df = pd.read_csv('SJM.csv', delimiter=',',usecols=['Number', 'Short description', 'Assignment group'])
    #identify esc/res tickets & place in to dataframes
    Helpdesk = 'IS 4th Source Helpdesk'
    Esc = df[(df['Assignment group'] != Helpdesk)]
    Res = df[(df['Assignment group'] == Helpdesk)]
    #convert esc/res dataframes to HTML
    Esc.to_html('SJM_Escalated.html')
    Res.to_html('SJM_Resolved.html')
    #open newly created html tables
    webbrowser.open_new_tab('SJM_Escalated.html')
    webbrowser.open_new_tab('SJM_Resolved.html')

def SJHD_Cat(imageName):
    df = pd.read_csv('SJHD.csv', delimiter=',')
    df = df.groupby(['Category']).size()
    #format 2-D bar chart
    df.plot(kind='pie', fontsize=10)
    plt.title('St. Jude Help Desk\nTickets by Category')
    plt.pie(df, autopct='%1.1f%%')
    plt.axis("equal")
    plt.xlabel('')
    plt.ylabel('')
    plt.savefig("File_{}.png".format(imageName))
    plt.show()
    #plt.subplots_adjust(left=None, bottom=.15, right=.15, top=.80, wspace=None, hspace=None)

def SJHD_SubCat(imageName):
    df = pd.read_csv('SJHD.csv', delimiter=',')
    df = df.groupby(['Subcategory']).size()
    #configure 2-D bar chart
    df.plot(kind='pie', fontsize=10)
    plt.title('St. Jude Help Desk\nTickets by Subcategory')
    plt.pie(df, autopct='%1.1f%%')
    plt.axis("equal")
    plt.xlabel('')
    plt.ylabel('')
    plt.savefig("File_{}.png".format(imageName))
    plt.show()
    #plt.subplots_adjust(left=None, bottom=.15, right=.15, top=.80, wspace=None, hspace=None)

def main():
    #create window object
    window = Tk()
    window.title('St. Jude Reporter-Generator-Automator v0.3')

    def run(command):
        (str(command))

    #combine functions via iteration
    def combine_funcs(*funcs):
        def combined_func(*args, **kwargs):
            for f in funcs:
                f(*args, **kwargs)
        return combined_func
    
    #menu
    L1 = Label(window, text="Step1.")
    L1.grid(row=0,column=0)
    L1 = Label(window, text="excel files to\nCSV format.")
    L1.grid(row=1,column=1)
    b1 = Button(window,text="Convert:", width=12, height=2, bg='lightblue', fg='black', command=convert_toCSV)
    b1.grid(row=1,column=0)

    #define SJHD buttons
    b1 = Button(window,text="Generate:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(combine_funcs(EscRes('SJHD.csv', 'St. Jude Helpdesk\nEscalated & Resolved Graph', 'SJHD_1')),
        (DailyTicketCount('SJHD.csv', 'SJHD_2')),
        (AsgmtGroup('SJHD.csv', 'SJHD_3')),
        (SJHD_TickReports()),
        (CallStat1('SJ Help Desk - Previous Week.csv',
            'St. Jude Help Desk\nCall Statistics\n',
            'SJHD_5')),
        (CallStat2(8, 'SJ Help Desk - Previous Week.csv',
            'St. Jude Help Desk\nCall Statistics (Averages)\n',
            ['Answer Speed', 'After Call Work Time', 'Call  Time', 'Handle Time', 'Hold Time', 'Queue Wait Time', 'Ring Time', 'Talk Time'],
            'SJHD_6')),
        (SJHD_Cat('SJHD_7')),
        (SJHD_SubCat('SJHD_8'))))
    #position button on grid
    b1.grid(row=1,column=2)

    b1 = Button(window, text = "Generate", width=12, height=2, bg='cyan', fg='black', command=lambda: run(combine_funcs(EscRes('SJM.csv', 'St. Jude Milli\nEscalated & Resolved Graph', 'SJM_1')),
        (DailyTicketCount('SJM.csv', 'SJM_2'),
        (AsgmtGroup('SJM.csv', 'SJM_3'),
        (SJM_TickReports()),
        (CallStat1('SJ Milli - Previous Week.csv',
            'St. Jude Milli Support\nCall Statistics\n',
            'SJM_4'),
        (CallStat2(7,'SJ Milli - Previous Week.csv',
            'St. Jude Milli Support\nCall Statistics (Averages)\n',
            ['Answer Speed', 'After Call Work Time', 'Call  Time', 'Hold Time', 'Queue Wait Time', 'Ring Time', 'Talk Time'],
            'SJM_5')))))))
    #position button on grid
    b1.grid(row=1,column=4)

    #define SJHD labels
    L1 = Label(window, text="St. Jude Helpdesk")
    L1.grid(row=0,column=3)
    L1 = Label(window, text="Escalated & Resolved 2-D Bar Graph")
    L1.grid(row=1,column=3)
    L1 = Label(window, text="Daily Ticket Breakdown 2-D Line Graph")
    L1.grid(row=2,column=3)
    L1 = Label(window, text="Active Assignment Group 2-D Bar Graph")
    L1.grid(row=3,column=3)
    L1 = Label(window, text="Escalated & Resolved HTML Reports")
    L1.grid(row=4,column=3)
    L1 = Label(window, text="Call Statistics 2-D Bar Graph")
    L1.grid(row=5,column=3)
    L1 = Label(window, text="Call Statistics Averages 2-D Bar Graph")
    L1.grid(row=6,column=3)
    L1 = Label(window, text="Tickets by Category Pie Chart")
    L1.grid(row=7,column=3)
    L1 = Label(window, text="Tickets by Subcategory Pie Chart")
    L1.grid(row=8,column=3)

    #define SJM labels
    L1 = Label(window, text="St. Jude Milli")
    L1.grid(row=0,column=5)
    L1 = Label(window, text="Escalated & Resolved 2-D Bar Graph")
    L1.grid(row=1,column=5)
    L1 = Label(window, text="Daily Ticket Breakdown 2-D Line Graph")
    L1.grid(row=2,column=5)
    L1 = Label(window, text="Active Assignment Group 2-D Bar Graph")
    L1.grid(row=3,column=5)
    L1 = Label(window, text="Escalated & Resolved HTML Reports")
    L1.grid(row=4,column=5)
    L1 = Label(window, text="Call Statistics 2-D Bar Graph")
    L1.grid(row=5,column=5)
    L1 = Label(window, text="Call Statistics Averages 2-D Bar Graph")
    L1.grid(row=6,column=5)

    window.mainloop()

main()
