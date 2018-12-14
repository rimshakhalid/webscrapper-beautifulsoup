from bs4 import BeautifulSoup
import xlsxwriter
import re

def clean_string(text):
    '''
    Take string as input and remove extra spaces from it
    '''
    return re.sub(' +'," ",text)
        
def create_xlsx(filename, dataset):
    '''
    Input:
    filename: string
    dataset: list
    Return:
    Create Excel file for data extracted
    '''
    print("Creating excel file...")
    filename = '{0}.xlsx'.format(filename)


    workbook = xlsxwriter.Workbook(filename) ## initializing excel object
    worksheet = workbook.add_worksheet() ## adding worksheet in excel


    ## excel settings
    row = col = 0 ## initializing rows and columns
    excel_header = ["Name", "Address", "Phone Number", "Type","Review"]
    bold = workbook.add_format({'bold': True, 'center_across': True})  ## Header settings

    ## adding header to excel
    for j, h in enumerate(excel_header):
        worksheet.write(row, col + j, h, bold)

    # adding data in excel
    row += 1
    for obj in dataset:
        worksheet.write(row, 0, clean_string(obj["name"]))
        worksheet.write(row, 1, clean_string(obj["address"]))
        worksheet.write(row, 2, clean_string(obj["phone"]))
        worksheet.write(row, 3, clean_string(obj["type"]))
        review = "{0} ({1})".format(clean_string(obj["review"]), clean_string(obj["stars"]))
        worksheet.write(row, 4, review)
        row += 1
        
    ## excel settings
    worksheet.freeze_panes(1, 0) ## freezing header so you can see column names when scroll down
    workbook.close() ## closing excel
    print("{} is successfully created".format(filename))
    


def parse_html():
    dataset = [] ## empty list to store each row object
    
    ## if you are using python 2.7 use following to read file
    html = open("yelp_listing.html", 'r')
    ## if you are using python 3 use following to read file
    ##html = open("yelp_listing.html", 'r', encoding="utf8")

    
    print("Parsing HTML file...")
    soup = BeautifulSoup(html,"lxml") ## Initializing beautiful soup with html text and use lxml parser


    ## Getting search result listing
    listings = soup.find_all('li' , {'class' : 'domtags--li__373c0__3TKyB list-item__373c0__M7vhU'})
    ## looping on <li> tags to extract data
    for li in listings:
        li_div = li.find('div', {'class':'businessName__373c0__1fTgn'}) ## getting div which contains title/name of restaurant
        ## checking whether current <li> elemnet is an Ad or not. Ignore <li> if it is an Ad
        if li_div and not li_div.find('p'):
            row = {} ## empty 

            ## Getting the name of restaurant
            row["name"] = ""
            if li_div.find('a'):
                row["name"] = li_div.find('a').text

            ## Getting the address of restaurant
            row["address"] = ""
            address = li.find('div', {'class':'domtags--div__373c0__3B6ae'})
            if address:
                row["address"] = address.text.replace("\n", " ").strip()
            elif li.find('address'):
                row["address"] = li.find("address").text.replace("\n", " ").strip()

            ## Getting the phone# of restaurant
            row["phone"] = ""
            phone = li.find('div', {'class':'lemon--div__373c0__6Tkil display--inline-block__373c0__2de_K border-color--default__373c0__2oFDT'})
            if phone:
                row["phone"] = phone.text.replace("\n", " ").strip()

            ## Getting the type/category of restaurant
            row["type"] = ""
            category = li.find('div', {'class':'priceCategory__373c0__3zW0R'})
            if category:
                row["type"] = category.text.replace("\n", " ").replace("$", "").strip()

            ## Getting the reviews of restaurant
            row["review"] = ""
            review = li.find('span' , {'class' : 'reviewCount__373c0__2r4xT'})
            if review:
                row["review"] = review.text.replace("\n", "").strip()

            ## Getting the stars of restaurant
            row["stars"] = ""
            stars = li.find('div' , {'class' : 'i-stars__373c0__Y2F3O'})
            if stars:
                row["stars"] = stars['aria-label'].replace("\n", "").strip()
            dataset.append(row)
            
    html.close()        
    return dataset


if __name__ == '__main__':

    data = parse_html()
    filename = "Data Scraped"
    create_xlsx(filename, data)
        

