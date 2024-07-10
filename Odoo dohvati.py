import selenium
import time
link= "https://euro-odoo.multinorm.local/web#id=8805&view_type=form&model=mrp.resource.planning&action=461"

tablica xpath = '//*[@id='notebook_page_14']/div/div[2]/div/table/tbody'

gumb xpath= '//*[@id='notebook_page_14']/div/div[2]/div/table/tbody/tr[1]/td[10]' # za svaki
projketi tab= '/html/body/div[1]/div/div[2]/div/div/div/div/div/ul/li[2]/a'
excel export gumb= '//*[@id="button_export_excel"]' # zatim da spremi download u temp mapu c://temp
