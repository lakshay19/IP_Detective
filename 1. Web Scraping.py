import os
import re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

options = webdriver.ChromeOptions()
options.add_argument("headless")
driver = webdriver.Chrome(executable_path='chromedriver', options=options)

df = pd.read_csv('input.csv')

# initialize all arrays for 1st dataset
all_urls = []
out_put_patent_number = []
out_put_priority_date = []
out_put_drawing = []
out_put_detail_des = []
out_put_abstract = []
fetched_patent = []

# claims
out_put_claim_set_number = []
out_put_claim_patent_number = []
out_put_claim_set = []
glob_pub_num = "US20070010915A1"
ignore_records = ["US6461380B1", "US20050088112A1", "US10993642B2", "WO2012078593A2", "US20070010915A1"]


def create_output_csv():
    output_data_frame = pd.DataFrame(data={"Patent_Number": out_put_patent_number,
                                           "Priority_Date": out_put_priority_date,
                                           "Abstract": out_put_abstract,
                                           "Disclosure": out_put_detail_des,
                                           "Figures": out_put_drawing})

    output_data_frame.to_excel("plaintiff_descriptions.xlsx", index=True)


def save_claims_csv(claim_set, pub_num_text, headers):
    print("Total Claims: " + str(len(claim_set)))
    for each in range(len(claim_set)):
        out_put_claim_set_number.append(each + 1)
        out_put_claim_patent_number.append(pub_num_text)
        out_put_claim_set.append(claim_set[each])

    output_data_frame = pd.DataFrame(data={"Patent_Number": out_put_claim_patent_number,
                                           "Claim_Set_Number": out_put_claim_set_number,
                                           "Claim_Set": out_put_claim_set})

    output_data_frame.to_excel("plaitiff_claims.xlsx", index=True)


def check_exists_by_xpath(xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True


def safe_to_text_file(text_file, des):
    file1 = open(text_file, "w", encoding="utf-8")
    file1.write(des)
    file1.close()


def init(ls):
    global glob_pub_num
    for i in ls:
        save_this_record = True
        # url = "https://patents.google.com/patent/US" + i
        url = "https://patents.google.com/patent/" + i
        if url not in all_urls and "RE" not in i and i not in ignore_records:
            glob_pub_num = i
            driver.get(url)
            all_urls.append(url)
            delay = 3
            print("URL: " + str(url))

            abs_text = ""
            string_detail_des = ""
            fetch_priority = ""
            figure_file_name = "None"
            disclosure_file_name = "None"

            # getting publication number
            pub_num = driver.find_element(By.XPATH, "//h2[@id='pubnum']")
            pub_num_text = pub_num.text

            # getting priority
            priority_div = driver.find_element(By.XPATH, "//div[@class='wrap style-scope application-timeline']")
            inner_html_priority = priority_div.get_attribute('innerHTML')
            priority_soup = BeautifulSoup(inner_html_priority, 'html.parser')
            priority_row = priority_soup.find_all("div",
                                                  {"class": "event layout horizontal style-scope application-timeline"})

            for row in priority_row:
                text = row.find("span", {"class": "title-text style-scope application-timeline"})
                if "Priority" in text.getText():
                    date = row.find("div", {"class": "priority style-scope application-timeline"})
                    try:
                        fetch_priority = date.getText()
                    except:
                        pass

            # getting claims
            if check_exists_by_xpath("//div[@class='claims style-scope patent-text']"):
                claim_set = []
                claim_div = driver.find_element(By.XPATH, "//div[@class='claims style-scope patent-text']")
                inner_html_claim = claim_div.get_attribute('innerHTML')
                claim_soup = BeautifulSoup(inner_html_claim, 'html.parser')

                array_grey = []
                array_grey_num = []
                grey_div = claim_soup.find_all("div", {"class": "claim-dependent style-scope patent-text"})
                for grey in grey_div:
                    value = grey.getText()
                    num = value.split('. ')[0]
                    try:
                        array_grey_num.append(int(num))
                        array_grey.append(value)
                    except:
                        save_this_record = False
                        print("-------- Ignore bad HTML for ["+str(glob_pub_num)+"]-------------")
                        continue

                if save_this_record:
                    all_divs = claim_soup.find_all("div", {"class": "claim-text style-scope patent-text"})
                    total_divs = len(all_divs)
                    str_claims = ""
                    counter = 0
                    for each in range(total_divs):
                        value = all_divs[each].getText()
                        try:
                            num = value.split('. ')[0]
                            int_num = int(num)
                            counter += 1
                            if int_num == 1:
                                str_claims += value
                            elif int_num > 1 and int_num not in array_grey_num:
                                claim_set.append(str_claims)
                                str_claims = ""
                                str_claims += value
                                # print(str(counter)+" - "+str(int_num))
                            elif int_num > 1 and int_num in array_grey_num:
                                str_claims += value
                        except:
                            str_claims += value

                    claim_set.append(str_claims)
                    save_claims_csv(claim_set, pub_num_text, False)

            # getting details, abstract and figure
            if check_exists_by_xpath("//div[@class='abstract style-scope patent-text']"):
                # abstract
                abstract = WebDriverWait(driver, delay).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='abstract style-scope patent-text']")))
                abs_text = abstract.text

                # details
                details_div = driver.find_element(By.XPATH, "//section[@id='description']")
                inner_html_details = details_div.get_attribute('innerHTML')
                details_soup = BeautifulSoup(inner_html_details, 'html.parser')
                before = ""
                all_of_the_text = details_soup.getText()
                all_tags = [tag for tag in details_soup.find_all()]
                for tag in all_tags:
                    tag_name = tag.name
                    value = tag.getText()
                    if tag_name == "heading":
                        if "DESCRIPTION OF THE DRAWING" in value:
                            before = value

                text = all_of_the_text[:all_of_the_text.index(before)]
                string_detail_des = os.linesep.join([s for s in text.splitlines() if s])

                if save_this_record:
                    if before != "":
                        string_drawing = all_of_the_text.split(before, 1)[1]
                        string_drawing = before + "\n" + string_drawing
                        figure_file_name = "output/"+i + "_figure.txt"
                        safe_to_text_file(figure_file_name, str(string_drawing))
                    else:
                        string_detail_des = os.linesep.join([s for s in all_of_the_text.splitlines() if s])

                    if string_detail_des != "":
                        disclosure_file_name = "output/" + i + "_disclosure.txt"
                        safe_to_text_file(disclosure_file_name, str(string_detail_des))

            # appending all extracted values in arrays
            if save_this_record:
                out_put_priority_date.append(fetch_priority)
                out_put_abstract.append(abs_text)
                out_put_patent_number.append(pub_num_text)
                out_put_detail_des.append(disclosure_file_name)
                out_put_drawing.append(figure_file_name)
                fetched_patent.append(i)

                create_output_csv()
                print("Publication Number: "+str(pub_num_text))
                print("Patent Number: "+str(i))
                print('-------------------')
        else:
            print("SKIP: "+i)

    print("All input records completed. Out file name: output.xlsx")


# print(check_exists_by_xpath("//div[@class='abstract style-scope patent-text']"))


try:
    input("Press Enter to execute the script.")
    all_patent = df['Publication_Number']
    list_patent = []
    append = False
    for each in all_patent:
        if each == glob_pub_num:
            append = True

        if append:
            list_patent.append(each)

    init(list_patent)
except:
    # WO2012078593A2 after this
    print("-------- Ignore bad HTML for ["+str(glob_pub_num)+"]-------------")
    ignore_records.append(glob_pub_num)
    all_patent = df['Publication_Number']
    list_patent = []
    append = False
    for each in all_patent:
        if each == glob_pub_num:
            append = True

        if append:
            list_patent.append(each)

    init(list_patent)

