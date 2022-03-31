# Python 3.7
import sys

if sys.version_info < (3,):
    print('Must use Python version 3')
    sys.exit()

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import win32com.client as win32
import os

browser = webdriver.Firefox()

def DownloadTabs(artist):
    artistFolder = os.path.join(os.getcwd(), 'Tabs', artist)

    if not os.path.exists(artistFolder):
        os.mkdir(artistFolder)

    artistFirstLetter = artist[0].lower()
    url = "https://www.guitartabs.cc/tabs/{}/{}".format(artistFirstLetter, artist)

    try:
        browser.get(url)
    except Exception as e:
        print(e)
        sys.exit()

    try:
        artistPage = browser.find_element_by_xpath("/html/body/div/table[2]/tbody/tr[2]/td[2]/div[2]/div[1]/h1") #h1
        if artistPage.text[:3] == '404':
            print('*No tabs found.')
            print(' Make sure the artist is spelled correctly and spaces are replaced with underscores.')
            print(' There might not be tabs for this artist.')
            return
    except:
        try:
            artistPage = browser.find_element_by_xpath("/html/body/div/table[2]/tbody/tr[2]/td[2]/div[2]/div[1]/span/h2")
            artistName = artist.replace('_', ' ')
            if artistPage.text[:len(artistName)] == artistName:
                pass
            else:
                print('*No tabs found.')
                print(' Make sure the artist is spelled correctly and spaces are replaced with underscores.')
                print(' There might not be tabs for this artist.')
                return
        except Exception as e:
            print('*No tabs found.')
            print(' Make sure the artist is spelled correctly and spaces are replaced with underscores.')
            print(' There might not be tabs for this artist.')
            return

    Pages = []
    Pages.append(url)

    pageIndex = 1
    while True:
        try:
            elem = browser.find_element_by_xpath("/html/body/div/table[2]/tbody/tr[2]/td[2]/div[2]/div[3]/a[{}]".format(pageIndex))
            if elem.get_attribute("href") not in Pages:
                Pages.append(elem.get_attribute("href"))
            pageIndex += 1
        except:
            break

    Pages = list(dict.fromkeys(Pages))

    tabLinks = []
    for page in Pages:
        browser.get(page)
        elems = browser.find_elements_by_class_name("ryzh2")
        for elem in elems:
            tabLinks.append(elem.get_attribute("href"))

    tabCount = len(tabLinks)
    max_song_len = max([len(link.split("/")[-1].replace(".html", "")) for link in tabLinks])

    if (max_song_len % 2) == 1:
        header = ' {0}#{0} |{1}Song{2}|  Exists  |  New Tab  |  Failed'.format(' ' * (int(len(str(max_song_len))) + 1), ' ' * int(round(max_song_len/2) - 2), ' ' * int(round(max_song_len/2) - 3))
    else:
        header = ' {0}#{0} |{1}Song{1}|  Exists  |  New Tab  |  Failed'.format(' ' * (int(len(str(max_song_len)))), ' ' * (int(max_song_len/2) - 2))
    print(header)
    print('-'*len(header))

    tabIndex = 1
    newTabs = 0
    for link in tabLinks:
        tabExists = False
        failed = False
        url_path = link.split("/")
        artist = url_path[-2]
        artist = artist.upper()
        song = url_path[-1]
        filename = '{}-{}'.format(artist, song)
        filepath = os.path.join(artistFolder, filename)

        if os.path.exists(filepath):
            tabExists = True
        else:
            browser.get(link)
            try:
                element = browser.find_element_by_xpath("/html/body/div[2]/table[2]/tbody/tr[2]/td[2]/div[2]/div[2]/div/div[4]")
                elementHTML = element.get_attribute('innerHTML')
            except:
                failed = True

            try:
                with open(filepath, 'w') as f:
                    f.write(elementHTML)
            except:
                failed = True

        if tabCount < 10:
            indexStr = '{}/{}'.format(tabIndex, tabCount)
        elif tabCount >= 10 and tabCount < 100:
            if tabIndex < 10:
                indexStr = ' {}/{}'.format(tabIndex, tabCount)
            else:
                indexStr = '{}/{}'.format(tabIndex, tabCount)
        elif tabCount >= 100 and tabCount < 1000:
            if tabIndex < 10:
                indexStr = '  {}/{}'.format(tabIndex, tabCount)
            elif tabIndex >= 10 and tabIndex < 100:
                indexStr = ' {}/{}'.format(tabIndex, tabCount)
            else:
                indexStr = '{}/{}'.format(tabIndex, tabCount)
        else:
            if tabIndex < 10:
                indexStr = '   {}/{}'.format(tabIndex, tabCount)
            elif tabIndex >= 10 and tabIndex < 100:
                indexStr = '  {}/{}'.format(tabIndex, tabCount)
            elif tabIndex >= 100 and tabIndex < 1000:
                indexStr = ' {}/{}'.format(tabIndex, tabCount)
            else:
                indexStr = '{}/{}'.format(tabIndex, tabCount)

        if tabExists:
            print(' {} |{}{}|    x     |  {}  |  {}'.format(indexStr, song.replace('.html', ''), ' ' * (max_song_len - len(song.replace('.html', ''))), ' ' * len('New Tab'), ' ' * len('Failed')))
        else:
            if failed:
                print(' {} |{}{}|  {}  |     x     |     x'.format(indexStr, song.replace('.html', ''), ' ' * (max_song_len - len(song.replace('.html', ''))), ' ' * len('Exists')))
            else:
                print(' {} |{}{}|  {}  |     x     |'.format(indexStr, song.replace('.html', ''), ' ' * (max_song_len - len(song.replace('.html', ''))), ' ' * len('Exists')))
                newTabs += 1
        tabIndex += 1

    return newTabs

def sendEmail(mail_to, tab_dict):
    print("\nFinished.")
    print("Sending notification email to {}".format(mail_to))
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mail_to

    mail.Subject = 'Guitar-Tab-Scrape.py has finished'
    mail.Body = 'New tabs downloaded: {}'.format(tab_dict)

    mail.Send()

def main(mail_to):
    Artists = ['Megadeth',
               'Boston',
               'As_I_Lay_Dying',
               'Trivium',
               'Pantera',
               'Judas_Priest',
               'Metallica',
               'Nevermore',
               'Lamb_Of_God',
               'Iron_Maiden',
               'All_That_Remains']

    if not os.path.exists(os.path.join(os.getcwd(), 'Tabs')):
        print('Output directory: {}\n'.format(os.path.join(os.getcwd(), 'Tabs')))
        os.mkdir(os.path.join(os.getcwd(), 'Tabs'))

    artistIndex = 1
    tabs = {}
    for artist in Artists:
        print("\n")
        print("*"*75)
        print('{}/{} - {}'.format(artistIndex, len(Artists), artist))
        print("*"*75)
        tabCount = DownloadTabs(artist)
        tabs.update({artist: tabCount})
        artistIndex += 1
    browser.close()
    sendEmail(mail_to, tabs)

if __name__ == "__main__":
    email = "grantcaleb22@gmail.com"
    print('')
    main(email)
