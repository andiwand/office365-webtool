import time

from selenium import webdriver

def get_name(element):
    return element.find_elements_by_css_selector("._ph_k6 > span")[0].text

def index_name(elements, name):
    for i, element in enumerate(elements):
        if get_name(element) == name: return i
    return -1

def index_selected(elements):
    for i, element in enumerate(elements):
        if element.get_attribute("tabindex") == "0": return i
        if element.get_attribute("aria-selected") == "true": return i
    return -1

driver = webdriver.Chrome()
driver.maximize_window()
driver.implicitly_wait(20)

driver.get("https://portal.office365.com")

login_username = driver.find_element_by_id("cred_userid_inputtext")
login_password = driver.find_element_by_id("cred_password_inputtext")
login_button = driver.find_element_by_id("cred_sign_in_button")

login_username.send_keys("username")
login_password.send_keys("password")
time.sleep(1)
login_button.click()

start_people = driver.find_element_by_id("ShellPeople_link")

driver.get("https://outlook.office365.com/owa/?realm=logsol.at&exsvurl=1&ll-cc=1031&modurl=2")

element = driver.find_element_by_xpath("//div[@aria-label='Ihre Kontakte']")
element.find_element_by_tag_name("button").click()
element.click();

people_list = driver.find_element_by_css_selector("._ph_Y5 > div")
people_links = driver.find_elements_by_css_selector("._ph_45")[1].find_elements_by_css_selector("._ph_55")[-3]

#TODO use logger

start = time.time()
count = 0
unlinked = 0
while True:
    #print "get cached contacts..."
    try:
        people_current = people_list.find_elements_by_css_selector("._ph_06")
        last = index_selected(people_current)
        if last < 0:
            print "exception. cannot find last. starting from first cached..."
            last = 0
        else:
            last += 1
    except:
        print "exception. people updated while searching. retry..."
        time.sleep(1)
        continue
    #print "found %d contacts, %d new" % (len(people_current), len(people_current) - last)
    people_current = people_current[last:]
    if len(people_current) <= 0: break
    for current in people_current:
        count += 1
        try:
            current.click()
        except:
            print "retry select..."
            time.sleep(1)
            break
        name = get_name(current)
        if people_links.value_of_css_property("display") != "none":
            print name
            while True:
                try:
                    time.sleep(1)
                    if people_links.value_of_css_property("display") == "none": break
                    people_links.click()
                    other_name = driver.find_elements_by_class_name("_pe_l2")[0].text
                    remove = driver.find_elements_by_class_name("_pe_o2")[0]
                    remove.click()
                    print "unlinked " + other_name
                    unlinked += 1
                except:
                    print "exception. nothing to remove (maybe double check by hand)."
                    break
    time.sleep(2)
end = time.time()

print "%d contacts in %d seconds" % (count, end - start)
print "%d unlinked" % unlinked

driver.close()
