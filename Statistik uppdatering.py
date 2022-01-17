import pandas as pd
import win32com.client
import datetime
import os.path
import time

def mail_svarstid():

    week_ago_sent = (datetime.datetime.now() - datetime.timedelta(days = 7)).strftime("%Y-%m-%d")
    week2_ago_sent = (datetime.datetime.now() - datetime.timedelta(days = 14)).strftime("%Y-%m-%d")
    
    handled_date_lower = (datetime.datetime.now() - datetime.timedelta(days = 30)).strftime("%Y-%m-%d")
    handled_date_upper = (datetime.datetime.now() - datetime.timedelta(days = 7)).strftime("%Y-%m-%d")
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    handled_folder = outlook.Folders("work@email.com").Folders("Inkorgen").Folders("Åtgärdade mail").Folders("Tim")
    sent_folder = outlook.Folders("work@email.com").Folders("Skickat")
    
    date_filter_sent = f"[SentOn] <= '{week_ago_sent}' and [SentOn] >= '{week2_ago_sent}'"
    date_filter_handled = f"[ReceivedTime] <= '{handled_date_upper}' and [ReceivedTime] >= '{handled_date_lower}'"
    
    conv_index_handled=[]
    handled_date=[]
    conv_index_sent=[]
    sent_date=[]
    conv_sent_id=[]
    conv_handled_id=[]
    for olItems_sent in sent_folder.Items.Restrict(date_filter_sent):
        if olItems_sent.Class == 43:
            
            conv_index_sent.append(olItems_sent.ConversationIndex)
            conv_sent_id.append(olItems_sent.ConversationID)
            sent_date.append(olItems_sent.ReceivedTime.strftime("%Y-%m-%d, %H:%M:%S"))
            
            
    
    for olItems_handled in handled_folder.Items.Restrict(date_filter_handled):
        if olItems_handled.Class == 43:
            
            conv_index_handled.append(olItems_handled.ConversationIndex)
            conv_handled_id.append(olItems_handled.ConversationID)
            handled_date.append(olItems_handled.SentOn.strftime("%Y-%m-%d, %H:%M:%S"))
            
    df_mail_sent = pd.DataFrame(list(zip(conv_index_sent, conv_sent_id, sent_date)), columns = ["Index", "ID", "Skickat Datum"])
    df_mail_handled = pd.DataFrame(list(zip(conv_index_handled, conv_handled_id, handled_date)), columns = ["Index", "ID", "Inkommet Datum"])
    df_response_time = pd.merge(df_mail_sent, df_mail_handled, on = 'ID')
    time_diff = pd.to_datetime(df_response_time["Skickat Datum"]) - pd.to_datetime(df_response_time["Inkommet Datum"])
    df_response_time['Svarstid'] = abs(time_diff)
    df_response_time.to_excel(r'H:\Statistik\replied.xlsx', index = False)



#Funktion för att köra mail statistik uppdatering VBA makro, den loggar antal inkomna mejl under dagen och totala olästa mejl under dagen----------------------------------------------------------------------------------------------
def mail_statistik_logg():
    
    xlApp = win32com.client.DispatchEx('Excel.Application')
    wb_mail_statistik = xlApp.Workbooks.Open(Filename=os.path.expanduser(r'H:\Statistik\Uppdatering.xlsm'))
    xlApp.Run('uppdatering')
    wb_mail_statistik.Save()
    wb_mail_statistik.Close(True)
    xlApp.Quit()
    del xlApp

def mail_aktuell_uppd():
    
    today_date = datetime.datetime.today().strftime("%Y-%m-%d")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    timavr_folder = outlook.Folders("work@email.com").Folders("Inkorgen").Folders("Timavräkning")
    adm_serier_folder = outlook.Folders("work@email.com").Folders("Inkorgen").Folders("Administrera serier")
    
    
    mail_date=[]
    title=[]
    cat_unread=[]
    adm_serier=[]
    
    for olItems_adm in adm_serier_folder.Items.Restrict("[Unread]=True"):
        if olItems_adm.SentOn.strftime('%Y-%m-%d') < today_date:
            adm_serier.append(olItems_adm.SentOn.strftime('%Y-%m-%d'))
            
    for olItems in timavr_folder.Items.Restrict("[Unread]=True"):
        if olItems.Categories == '':
            olItems.Categories = 'Okategoriserad'
        cat_unread.append(olItems.Categories)
        title.append(olItems.Subject)
        mail_date.append(olItems.SentOn.strftime('%Y-%m-%d'))
    
    df_mail_unread = pd.DataFrame(list(zip(mail_date, cat_unread,title)), columns = ["Datum", "Kategori", "Rubrik"])
    df_adm_serier_manual = pd.DataFrame(adm_serier, columns = ["Inkommet Datum"])
    df_adm_serier_manual = df_adm_serier_manual.groupby(by=["Inkommet Datum"])["Inkommet Datum"].count().to_frame(name = "Antal Manuella Admin serier").reset_index()
    matfel_folder = outlook.Folders("work@email.com").Folders("Inkorgen").Folders("Timavräkning").Folders("Mätfel & Bilaterala frågor")

    unread_matfel_mail_title =[]
    unread_matfel_mail_date = []
    
    for olItems_unread_matfel in matfel_folder.Items.Restrict("[Unread]=True"):
        if olItems_unread_matfel.Class == 43:
            
            unread_matfel_mail_title.append(olItems_unread_matfel.Subject)
            unread_matfel_mail_date.append(olItems_unread_matfel.SentOn.strftime('%Y-%m-%d'))
    df_unread_bilaterala_matfel_mail = pd.DataFrame(list(zip(unread_matfel_mail_date, unread_matfel_mail_title)), columns=["Datum", "Rubrik"])
    df_unread_bilaterala_matfel_mail = df_unread_bilaterala_matfel_mail.groupby(by=["Datum"])["Datum"].count().to_frame(name = "Antal Mail per Inkommet Datum").reset_index()
    
    writer = pd.ExcelWriter(r'\\workspaces.office.com\DavWWWRoot\sites\settlement\Uppgifter i gruppen\Excelverktyg\Statistik\Aktuell Mail Info.xlsx', engine = 'xlsxwriter')
    df_mail_unread.to_excel(writer, sheet_name = 'Olästa Mail Kat',index = False)
    df_adm_serier_manual.to_excel(writer, sheet_name = 'Adm Serier Hantera Manuellt',index = False)
    df_unread_bilaterala_matfel_mail.to_excel(writer, sheet_name = 'Mail Mätfel & Bilaterala',index = False)
    writer.save()
    
    print("Mail Statistik Uppdaterad: " + str(datetime.datetime.now()))
    
pause_time = 1800

while datetime.datetime.now().hour < 18:
    if datetime.datetime.now().hour >= 8 and datetime.datetime.now().hour <= 12:
        mail_svarstid()
        mail_statistik_logg()
        mail_aktuell_uppd()
        print("1")
        time.sleep(pause_time)
        
    elif datetime.datetime.now().hour >= 14 and datetime.datetime.now().hour <= 16: 
        mail_statistik_logg()
        mail_aktuell_uppd()
        print("2")
        time.sleep(pause_time)
        
    elif datetime.datetime.now().hour >= 17:
        mail_statistik_logg()
        mail_aktuell_uppd()
        print("3")
        time.sleep(pause_time)
        
    elif datetime.datetime.now().hour >= 18:
        break
        
    else:
        print("paus")
        print(datetime.datetime.now())
        time.sleep(pause_time)
    
    print("paus avslutat")