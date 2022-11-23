# Write collected person info to excel format. Need to go faster- its almost 9pm.
#
# Author: Kyle T. / scritty
#
# Runtime notes: This script and the module table_data_parser needs the xlsxwriter
# module to properly save the information, and a working internet connection to get
# the table data for Monterey Peninsula College.
#
# xlsxwriter module can be installed with the following command: pip install xlsxwriter

import os
import requests
import xlsxwriter
from table_data_parser import collectContacts

schools = {
    "Los Medanos":
        {
            "contacts": [],
            "link": "https://www.losmedanos.edu",
            "comments": "",
        	"campus":"Pittsburg Campus",
			"postal":"94565",
		},
    "Madera": # Website undergoing "maintentence"- everything is either outdated, insufficient contact info, or not available
        {
            "contacts": [],
            "link": "https://www.maderacenter.com",
            "comments": "",
        	"campus":"",
			"postal":"",
		},
    # Maymount is closed, omitting...
    "Mendocino":
        {
            "contacts": [],
            "link": "https://www.mendocino.edu",
            "comments": "Specific campus locations were not listed in the Staff Directory alongside persons.",
        	"campus":"Ukiah Campus",
			"postal":"95482",
        },
    "Merced":
        {
            "contacts": [],
            "link": "https://www.mccd.edu",
            "comments": "",
        	"campus": "",
			"postal": "95348",
		},
    "Merritt":
        {
            "contacts": [],
            "link": "https://www.merritt.edu",
            "comments": "",
        	"campus": "",
			"postal": "94619",
		},
    "MiraCosta":
        {
            "contacts": [],
            "link": "https://www.miracosta.edu",
            "comments": "Extensions beginning with 66, 67, 68 or extensions 8700-8749 may be dialed direct using 760.795.xxxx\nExtensions beginning with 78 may be dialed direct using 760.634.xxxx\nExtensions 2012-2411 may be dialed direct using 442.262.xxxx\nAll other extensions may only be reached by dialing 760.757.2121 and entering the four-digit extension when prompted.",
        	"campus": "Oceanside Campus",
			"postal": "92056",
		},
    "Mission":
        {
            "contacts": [],
            "link": "https://missioncollege.edu",
            "comments": "",
        	"campus": "",
			"postal": "95054",
		},
    "Modesto Junior": # Needs to be done manually- the entire thing is stored in <div>'s... I cry
        {
            "contacts": [],
            "link": "https://www.mjc.edu",
            "comments": "",
        	"campus": "",
			"postal": "95350",
		},
    "Monterey Peninsula": # Horribly created website- every 75 table entries is stored in an entirely different webpage, and must be scraped individually
        {
            "contacts": [],
            "link": "https://www.mpc.edu",
            "comments": "",
            "campus": "Marina Campus",
			"postal": "93940",
        },
}

# Colleges that must be handled seperately
special_schools = ["Madera", "Merced", "Modesto Junior", "Monterey Peninsula"]

# Aggregating Data
for school in schools:
    if school not in special_schools:
        name = school + " College"
        file = "tables/" + name + ".txt"
        print("Collecting data for %s" % name)
        if os.path.exists(file):
            schools[school]['contacts'] = [collectContacts(file)]
            print("Done.")
        else:
            schools[school]['contacts'] = None
            print("Skipped missing file.")

# Handling "special" cases
for school in special_schools:
    if school == "Monterey Peninsula":
        
        # setting up to get the table elements via webrequests (im not copy pasting all of that by hand)
        headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36',
                'Content-Type': 'text/html',
                }
        
        print("Fetching table data for %s" % school)
        for num in range(1, 9):
            resp = requests.get("https://www.mpc.edu/about-mpc/campus-information/contact-us/faculty-staff-directory/-npage-%s" % num, headers=headers)
            table_element = resp.text[resp.text.find("<table"):resp.text.find("</table>") + len("</table>")]
            with open("tables/%s/npage-%s.txt" % (school, num), 'w') as file: file.write(table_element)
        print("Table saved.")
        
        print("Collecting data for %s College" % school)
        # parsing tables seperately
        for file in os.listdir("tables/" + school):
            schools[school]['contacts'].append(collectContacts("tables/%s/%s" % (school, file), dept_flip=True))
        print("Done.")
    
    if school == "Modesto Junior":
        
        # Parsing data manually from a handwritten text file
        print("Collecting data for %s College" % school)
        with open("tables/%s College.txt" % school, 'r') as file: data = file.read().splitlines()
        schools[school]['contacts'] = [
            [{
                'email': data[line + 2],
                'phone': data[line + 3],
                'first': data[line].split(' ')[0],
                'middle': '',
                'last': data[line].split(' ')[1],
                'department': data[line + 1].rsplit('-', 1)[-1],
                'location': ""
            } for line in range(0, len(data), 5)]
        ]
        print("Done.")
        
    if school == "Merced":
        
        # Add phone area code when collecting contact information
        name = school + " College"
        file = "tables/" + name + ".txt"
        print("Collecting data for %s" % name)
        if os.path.exists(file):
            schools[school]['contacts'] = [collectContacts(file, dept_flip=False, phone_ext="209")]
            print("Done.")
        else:
            schools[school]['contacts'] = None
            print("Skipped missing file.")

# Creating a spreadsheet
workbook = xlsxwriter.Workbook("CEDB-Data-CA08-TranK[BOT].xlsx")
sheet = workbook.add_worksheet()

# Creating formats
bold = workbook.add_format({'bold': True})

# Adding headers for readability
sheet.write("A1", "ResInit", bold) # Static: "KTT"
sheet.write("B1", "DateVerified", bold) # Static: "11/22/2022"
sheet.write("C1", "CollegeListTag", bold) # Static: "CA08"
sheet.write("D1", "CollegeName", bold)
sheet.write("E1", "CampusName", bold)
sheet.write("F1", "City", bold)
sheet.write("G1", "State", bold)
sheet.write("H1", "PostalCode", bold)
sheet.write("I1", "Country", bold) # Static: "USA"
sheet.write("J1", "Website", bold)
sheet.write("K1", "Department", bold)
sheet.write("L1", "BAorBS", bold)
sheet.write("M1", "FacultyLastName", bold)
sheet.write("N1", "FacultyFirstName", bold)
sheet.write("O1", "FacultyMiddleInital", bold)
sheet.write("P1", "FacultyEmail", bold)
sheet.write("Q1", "FacultyPhone", bold)
sheet.write("R1", "Comments", bold)
# ResInit | DateVerified | CollegeListTag | CollegeName | CampusName | City | State | PostalCode | Country
# Website | Department | BAorBS | FacultyLastName | FacultyFirstName | FacultyMiddleInitial | FacultyEmail | FacultyPhone | Comments


# Writing Programatic Data
initals = "KTT"
date = "11/22/2022"
state = "CA"
country = "USA"
tag = "CA08"
column = 2
for school in schools.items():
    name, info = school
    name += " College"

    if info['contacts']:
        for people in info['contacts']:
            for person in people:
            
                # Writing static data
                sheet.write("A%s" % column, initals) # ResInit
                sheet.write("B%s" % column, date) # DateVerified
                sheet.write("C%s" % column, tag) # CollegeListTag
                sheet.write("G%s" % column, state) # State
                sheet.write("I%s" % column, country) # Country
                
                # Writing school information
                sheet.write("D%s" % column, name) # College Name
                sheet.write("E%s" % column, info['campus']) # Preset campus name
                sheet.write("H%s" % column, info['postal']) # Preset postal code
                sheet.write("J%s" % column, info['link']) # College website
                sheet.write("R%s" % column, info['comments']) # Preset comments
                
                # Writing contact information
                sheet.write("K%s" % column, person['department']) # Department
                sheet.write("M%s" % column, person['last']) # FacultyLastName
                sheet.write("N%s" % column, person['first']) # FacultyFirstName
                sheet.write("O%s" % column, person['middle']) # FacultyMiddleInital
                sheet.write("P%s" % column, person['email']) # FacultyEmail
                sheet.write("Q%s" % column, person['phone']) # FacultyPhone
                
                # Finding specific location
                if name == "MiraCosta College":
                    if "CLC" in person['location']:
                        sheet.write("E%s" % column, "Community Learning Center")
                        sheet.write("H%s" % column, "92058")
                    elif 'SAN' in person['location']:
                        sheet.write("E%s" % column, "San Elijo Campus")
                        sheet.write("H%s" % column, "92007")
                    elif 'TCI' in person['location']:
                        sheet.write("E%s" % column, "Technology Career Institute & North San Diego Small Business Development Center")
                        sheet.write("H%s" % column, "92011")
                
                # Moving to next column
                column += 1
            
workbook.close()
print("Complete!")