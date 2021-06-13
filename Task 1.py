import time #using time library to give pauses while parsing to different pages.
from xlwt import Workbook #using xlwt library to create excel file and save data in it.
from bs4 import BeautifulSoup #using BeautifulSoup library to scrape data from webpage.
from selenium import webdriver #using selenium to open the webpage in chrome browser.


#Below are the urls of all the categories under computer tag on AMAZON
urls = ['https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A172456&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_0',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A193870011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_1',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A13896617011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_2',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A1292110011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_3',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A3011391011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_4',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A1292115011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_5',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A172504&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_6',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A17854127011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_7',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A172635&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_8',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A172584&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_9',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A11036071&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_10',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A2348628011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_11',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A15524379011&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_12',
           'https://www.amazon.com/s?bbn=16225007011&rh=n%3A16225007011%2Cn%3A16285851&dc&qid=1623583239&rnid=16225007011&ref=lp_16225007011_nr_n_13']
 
#Below are the categorie's names under computer tag on AMAZON
categories = ['Computer Accessories and Peripherals',
              'Computer Components',
              'Computers & Tablets','Data Storage',
              'Laptop Accessories',
              'Monitors',
              'Networking Products',
              'Power Strips & Surge Protectors',
              'Printers',
              'Scanners',
              'Servers',
              'Tablet Accessories',
              'Tablet Replacement Parts',
              'Warranties & Services']

#I used counter here, which will tell us how many pages we have scraped after the end of page.(working shown below)
counter = 1

# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet in excel file. 
sheet1 = wb.add_sheet('Sheet 1') 
# row means the current selected row for adding products
row = 1
# col means the current selected column for adding products
col = 1
# colm means the current selected column for add product categories 
colm = 0

#below 3 lines are used to add the headings of excel file
sheet1.write(0,0,'Product Categories')
sheet1.write(0,1,'Product\'s')
sheet1.write(0,2,'Badge')

#here we are initiating the chrome webdriver to parse the website.
wd = webdriver.Chrome('chromedriver')

# As you can see above, we have 14 categories, so we are creating a loop which will be repeated 14 times and will perform the actions come's under this loop
for contents in range(14):
    
    #below we are parsing the urls
    wd.get(urls[contents])
    #here we are storing the page_source in response variable
    response = wd.page_source
    # here we are using BeautifulSoup library to create the parsed tree which we will use to scrape data
    soup = BeautifulSoup(response, 'html.parser')
    
    #currently, I am scraping only first page of every category, for scraping all pages, replace below for loop with while loop
    #while(True): 
    for i in range(5):#here, 5 means this loop will only scrape first 5 pages, you can increase or decrease this number.
        
        #here, I have scraped the div class of the name mentioned below from the beautifulsoup parsed tree
        for data in soup.find_all('div', {'class': 's-expand-height s-include-content-margin s-border-bottom s-latency-cf-section'}):
            
            #same as above, I have scraped the span class for product name
            new_data = data.find('span',{'class': 'a-size-base-plus a-color-base a-text-normal'})
            
            #I used if statement here, which will check whether this current item have BEST SELLER badge or not
            #if there is no BEST SELLER badge then it will pass the excel cell and no value will be added to that cell
            if data.find('div', {'class': 'a-row a-badge-region'}):
                badge = data.find('div', {'class': 'a-row a-badge-region'})
                #the excel badge cell comes under span class of name 'a-badge-text'
                bs_badge = badge.find('span',{'class': 'a-badge-text'})
                
                #printing the badge
                print(bs_badge.text)
                #saving the badge in column c of excel file
                sheet1.write(row,2,bs_badge.text)
            else:
                #if the product have no badge then it will simple print NOT A BEST SELLER and no data will be added to excel file
                print("NOT A BEST SELLER")
                
            #here, I have only selected the text and not the whole HTML tag of the product name
            a = new_data.text
            
            #printing the product name
            print(a)
            
            #adding the product categories in excel file
            sheet1.write(row,colm,categories[contents])
            #after adding a product category I have moved to next row
            rowm += 1
            #adding the product name to the row,col 
            sheet1.write(row,col, a)
            #moving the next row after adding product on a specific cell
            row += 1
            #saving the excel file, here I am saving the excel file after every iteration as to avoid any data loss if internet goes offline or we found some error on webpage.
            wb.save('DVIZ Test.xlsx')
        
        #the below command will print the total no. of pages we have scrped on the website.
        print("Page's Scraped : "+ str(counter))
        #increasing the counter value so next time it shows the correct no. of scraped pages.
        counter += 1
        #here, I have used time.sleep function which will pause the whole process for 10 seconds, it is because if parse the AMAZON so quicky, then we out IP address might get blocked.
        #If you want to scrape large amount of data then the sleep time should be random, so in that case we have to add random sleep time after every iteration.
        time.sleep(10)
        
        #below command is used to go to the next page if the the NEXT button is enabled as on last page it becomed disabled.
        #We are using try except as well, because on last page the element('a-last') will not be present and our program will crash,
        #so to overcome crashing, I am using TRY EXCEPT
        try:
            if wd.find_element_by_class_name('a-last').is_enabled():
                wd.find_element_by_class_name('a-last').click()
            #the argument under else command is breaking the loop, this will be helpful if we are using while loop, because in every category we will have different no. of pages, so this will take care of that.
            else:
                break
        except Exception as e:
            #this is basically printing the exception, it not compulsory as we know that we will get error if we are on the last page of a specific product category.
            print(e)
wd.quit()