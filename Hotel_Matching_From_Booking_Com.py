import openpyxl
import time
from selenium import webdriver
import csv

# Give the path of the xlsx file from where it will take Hotel, city and country name.
path = "E:/PyCharm Projects/Input_File.xlsx"

# To open the workbook
# workbook object is created
wb_obj = openpyxl.load_workbook(path)

# Get workbook active sheet object
# from the active attribute
sheet_obj = wb_obj.active

# Cell objects also have row, column,
# and coordinate attributes that provide
# location information for the cell.

# Note: The first row or
# column integer is 1, not 0.

# Cell object is created by using
# sheet object's cell() method.
cell_obj = sheet_obj.cell(row=1, column=1)

# Declare lists for all inputs.
input_hotel_list = []
input_city_list = []
input_country_list = []

# Just initialize 2values in list for index starts from 2.
new_map_zone_urls_list = ['Url_Col', 'index_val']
# Find out maximum number of rows in the worksheet
max_row = sheet_obj.max_row
# Iterating values of a particular column number 3 i.e "New Mapping Zone Names" column.
for i in range(2, 12):#(max_row + 1)
    cell_obj = sheet_obj.cell(row=i, column=5)
    input_hotel_list.append(cell_obj.value)  # Take all hotels from input file
    cell_obj = sheet_obj.cell(row=i, column=7)
    input_city_list.append(cell_obj.value)  # Take all cities from input file
    cell_obj = sheet_obj.cell(row=i, column=6)
    input_country_list.append(cell_obj.value)  # Take all country from input file

# #Check all input lists
# for hotel, city, country in zip(input_hotel_list,input_city_list,input_country_list):
#     print("{}-{}-{}".format(hotel,city,country))

driver = webdriver.Chrome("E:/PyCharm Projects/chromedriver.exe")
# Iterating "New Mapping hotel Names" column values one by one.
for count_4_header, (hotel, city, country) in enumerate(zip(input_hotel_list, input_city_list, input_country_list)):
    # Hitting booking.com
    driver.get('https://www.booking.com/')
    time.sleep(3)
    # Find Hotel's input box element
    search_box = driver.find_element_by_xpath(
        '//input[@class="c-autocomplete__input sb-searchbox__input sb-destination__input"]')
    search_box.clear()
    # Pass search string(Hotel's name) on input box.
    search_box.send_keys(hotel)
    time.sleep(3)
    # ##################write a response file after passing a Hotel's name ####################
    Location_mapping_Resp = driver.page_source
    Loc_map = open("Location_mapping" + hotel + ".html", "w", encoding="utf-8")
    Loc_map.write(Location_mapping_Resp)
    ###########################################################
    # Storing suggested Hotel's element into a list.

    Hotel_list = driver.find_elements_by_xpath(
        '//li[@class="c-autocomplete__item sb-autocomplete__item sb-autocomplete__item-with_photo sb-autocomplete__item--hotel sb-autocomplete__item__item--elipsis "]/span/b[@class="search_hl_name"]')
    Hotel_list_with_address = driver.find_elements_by_xpath(
        '//li[@class="c-autocomplete__item sb-autocomplete__item sb-autocomplete__item-with_photo sb-autocomplete__item--hotel sb-autocomplete__item__item--elipsis "]/span')

    for hotel_name, Hotel_add in zip(Hotel_list, Hotel_list_with_address):
        if (city in Hotel_add.text) and (country in Hotel_add.text):

            with open("Output_Hotel_Matching.csv", "a", newline='') as file:
                # Defines column names into a csv file.
                field_names = ['Input_HotelName', 'Input_CityName', 'Input_CountryName', 'Output_HotelName',
                               'Output_HotelName_Address']
                writer = csv.DictWriter(file, fieldnames=field_names)
                # Condition for writing header only once.
                if count_4_header == 0:
                    writer.writeheader()
                count_4_header += 1
                # Writing all information in a row.
                writer.writerow(
                    {
                        'Input_HotelName': hotel,
                        'Input_CityName': city,
                        'Input_CountryName': country,
                        'Output_HotelName': hotel_name.text,
                        'Output_HotelName_Address': Hotel_add.text

                    }
                )
driver.quit()#Close brower
