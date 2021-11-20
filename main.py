import dateutil
import urllib.request
import tweepy
import pandas as pd

from credentials import *

opener = urllib.request.build_opener()
opener.addheaders = [('User-agent', 'Mozilla/5.0')]
urllib.request.install_opener(opener)
urllib.request.urlretrieve("https://www.rki.de/DE/Content/InfAZ/N/Neuartiges_Coronavirus/Daten/Impfquotenmonitoring.xlsx?__blob=publicationFile", 'Impfquotenmonitoring.xlsx')

xl = pd.ExcelFile('Impfquotenmonitoring.xlsx')
df = xl.parse('Impfungen_proTag')
df = df.dropna()
df = df[:-1] #drop last row
# print(df[1:])
#df = df[2:].astype(int)
df['Datum'] = pd.to_datetime(df['Datum'])
df['ZweitimpfungSum'] = df['Zweitimpfung'].cumsum()
df['AuffrischimpfungSum'] = df['Auffrischimpfung'].cumsum()

#print(df)

sixMonths = dateutil.relativedelta.relativedelta(months=6)

sixMonthsAgo = df['Datum'].iloc[-1]-sixMonths
rowSixMonthsAgo = df.loc[df['Datum'] == sixMonthsAgo]
fullyVaxxedSixMonthsAgo = rowSixMonthsAgo['ZweitimpfungSum'].values[0]
boostedNow = df['AuffrischimpfungSum'].iloc[-1] 

differenceOfShame = fullyVaxxedSixMonthsAgo - boostedNow

tweet = "Vor 6 Monaten waren {:,} Menschen vollständig geimpft, bis heute haben {:,} Menschen eine Auffrischung bekommen. Wir liegen {:,} Impfungen zurück.".format(int(fullyVaxxedSixMonthsAgo) , int(boostedNow) ,int(differenceOfShame))


print(tweet) 

# Authenticate to Twitter
auth = tweepy.OAuthHandler(consumer_key, consumer_secret_key)
auth.set_access_token(access_token, access_token_secret)
api = tweepy.API(auth)

try:
    api.verify_credentials()
    api.update_status(tweet)
    print("Authentication Successful")
except:
    print("Authentication Error")
