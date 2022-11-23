# This program is meant to read a text file of HTML table data
# and record the name and contact information of the specified fields:
#
# (Computer Engineering, Computer Science, Cybersecurity, Data Science, Information Systems, Information Technology, and Software Engineering)
#
# Author: Kyle T. / scritty

import re as regex
from html import unescape

search_terms = ["computer", "software", "information tech", "technology", "cyber"]
exclusion_terms = ["associate", "assistant", "intern"]

def search(iterable, search_terms, flip = False, verbose = False):
    # Loops down a list until it finds the specified search string, then returns the index
    if verbose: print("Search Terms: " + ', '.join(search_terms))
    if flip: iterable.reverse()
    for i, t in enumerate(iterable):
        if verbose: print("Searching [%s]: %s" % (i, t))
        for term in search_terms:
            # Regex filter out hypenated words
            if len(regex.findall(r'([^\-\"])%s(?!\=).*?' % term, t.lower())):
                if flip:
                    return list(range(0, len(iterable)))[::-1].index(i)
                else:
                    return i
    return None

def parse_row(row_element: str, name_column, seperate_name_column = None, dept_column = None, extension = None, verbose = False):
    """Takes in a table row (<tr>) element, and parses for a person's name and contact info

    Args:
        row_element (str): String representation of a table row element
        
    Returns:
        A dictionary value of the person's First, Middle, and Last names with their contact information.
    """
    person = {
        'email': '',
        'phone': '',
        'first': '',
        'middle': '',
        'last': '',
        'department': '',
        'location': ""
    }

    if seperate_name_column is None:
        name_str = row_element.split('/td>')[name_column]
        if verbose: print("Name RawString: %s" % name_str)
        name_str = regex.findall(r'\>(\s+)?([A-Z][\-a-zA-Z]+)(\s+)?(\S.*?)?(,)?(\s+)?([A-Z][\-a-zA-Z]+)(\s+)?\<', name_str)
        if verbose: print("Name FilterString: %s" % name_str)
        if len(name_str):
            name_str = [d.strip() for d in name_str[0] if d.strip() != '']
            if ',' in name_str:
                name_str.pop(name_str.index(','))
                name_str.reverse() # flip to Last, First name order
            if len(name_str) == 3:
                person['first'], person['middle'], person['last'] = name_str
            else:
                person['first'], person['last'] = name_str
    else:
        person['first'] = regex.findall(r'\>(\s+)?([a-zA-Z\-\. ]+)(\s+)?\<', row_element.split('/td>')[name_column])
        if verbose: print("First Name: %s" % person['first'])
        person['first'] = person['first'][0][1].strip()
        person['last'] = regex.findall(r'\>(\s+)?([a-zA-Z\-\. ]+)(\s+)?\<', row_element.split('/td>')[seperate_name_column])
        if verbose: print("Last Name: %s" % person['last'])
        person['last'] = person['last'][0][1].strip()
    
    if dept_column is not None:
        dept_str = row_element.split('/td>')[dept_column]
        if verbose: print("Dept. RawString: %s" % dept_str)
        dept_str = regex.findall(r'\>(\s+)?(?!\<)(\S.*?)(\s+)?\<', dept_str)
        if verbose: print("Dept. String: ", end='') or print(dept_str)
        if len(dept_str) and any([term in dept_str[0][1].lower() for term in search_terms]) and not any([term in dept_str[0][1].lower() for term in exclusion_terms]): person['department'] = dept_str[0][1]
    
    
    for line in row_element.splitlines():
        line = line.strip()
        
        # Getting email address
        email = regex.findall(r'[^\s\"\<]([a-zA-Z0-9\.]+?\@[^\s\"\>\<]+\.edu)', line)
        if len(email):
            if verbose: print("Email: ", end='') or print(email[0])
            person['email'] = email[0]
        
        # Getting phone number
        if extension is None:
            phone = regex.findall(r'(\s+)?([1])?(\s+)?(\()?([0-9][0-9][0-9])(\) |\s+|\-)([0-9][0-9][0-9]\-[0-9][0-9][0-9][0-9])', line)
            if len(phone):
                person['phone'] = "%s%s-%s" % (phone[0][2], phone[0][4], phone[0][6])
        else:
            phone = regex.findall(r'([0-9][0-9][0-9]\-[0-9][0-9][0-9][0-9])', line)
            if len(phone):
                person['phone'] = "%s-%s" % (extension, phone[0])
                
        # Storing location (room number)
        location = regex.findall(r'[^\s\"\<](\s+)?([A-Z][A-Z][A-Z0-9]+)(\s+)?[^\"\>]', line)
        if len(location):
            person['location'] = location[0][1]
        
    if person['email'] == '' or person['department'] == '' or person['first'].capitalize() != person['first']: return None
    # Cleaning up middle inital
    if person['middle'] != '':
        person['middle'] = regex.findall(r'([A-Z])', person['middle'])[0]
    
    return person

def collectContacts(filename: str, verbose: bool = False, dept_flip = False, phone_ext = None):
    """Collects the information of persons stored within an HTML Table element.

    Args:
        filename (str): Name of the text file containing the table.
        verbose (bool): Show the column indexes for manual inspection. Defaults to False.

    Returns:
        list[dict]: Holds 'email', 'phone', 'first', 'middle', 'last', and 'department'.
    """
    with open(filename, 'r', errors='ignore') as file: table = unescape(file.read())
        
    head = table.split('thead', 1)[1].split('</thead>', 1)[0]
    body = table.split('tbody', 1)[1].split('</tbody>', 1)[0]

    # finding the "name" column, if possible
    if "first" in head.lower() and "last" in head.lower(): # first and last names stored in seperate columns
        column = search(head.split('</th>'), ["first"])
        column_two = search(head.split('</th>'), ["last"])
    else:
        column = search(head.split('</th>'), ["name", "staff"])
        column_two = None
        
    # finding the "department" column
    dept_column = search(head.split('</th>'), ["dept", "department", "title"], dept_flip, verbose)

    # Showing determined columns for manual verification
    if verbose: print("Name Column: %s %s\nDepartment Column: %s" % (column, column_two, dept_column))


    # getting row data
    r = [t.split('</tr>')[0].strip() for t in body.split('<tr')]
    rows = []
    for row in r:
        data = '\n'.join([d.strip() for d in row.splitlines()])
        rows.append(data)

    # Filter for CS professions
    rows = [row for row in rows if any([term in row.lower() for term in search_terms])]
    contacts = [d for d in [parse_row(data, column, column_two, dept_column, phone_ext, verbose) for data in rows] if d is not None]
    return contacts