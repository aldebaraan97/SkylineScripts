#! python3
# PythonFile - Gets the texts for all projects in basecamp and
# sorts them alphabetically in an Excel spread sheet.
# Libraries needed to install: BeautifulSoup4, openpyxl
import bs4, openpyxl

htmlFile = "htmlFiles/https __3.basecamp.com_4019487_projects_directory view=active.html"
filename = "spreadSheets/basecampProjects.xlsx" 

# Converts html into a BeautifulSoup object
def pagesToString(page):
    pageToString = open(page, encoding="utf8")
    soup = bs4.BeautifulSoup(pageToString, "html.parser") 
    return soup

# Looks for an specific element in the BeautifulSoup object.
def getProjectList():
    tag = "a"
    attribute = "project-list__link list__filterable-content flex-1"
    soup = pagesToString(htmlFile)
    # Saves, sorts and cleans the strings.
    projectTitles = sorted([title.get("title").strip() for title in soup.find_all(tag, class_=attribute)], key=str.lower)
    return projectTitles

def getOldList():
    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active
        oldList = []
        for cellObj in list(sheet.columns)[0]:
            print("No data.") if cellObj.value == None else oldList.append(cellObj.value.strip())
        sorted(oldList, key=str.lower)
        return oldList
    except IndexError:
        oldList = []
        return oldList

# Compares lists.
def getUpDatedList():
    oldList = getOldList()
    newProjects = [project.strip() for project in getProjectList() if project not in oldList]  
    sorted(newProjects, key=str.lower)
    return newProjects

# Update and edit sheet.
def updateList():
    try:
        wb = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        wb.save(filename)
    sheet = wb.active

    oldList = getOldList()[:]
    updatedList = getUpDatedList()[:]

    # Edit sheet.
    sheet["A1"]= "Project Name"
    sheet["B1"]= "Archived"

    # Compares the two lists and adds new titles.
    for project in updatedList:
        if project not in oldList:
            print(f"New project to add: {project}")
            projectIndex = updatedList.index(project) + 2 
            sheet.insert_rows(idx=projectIndex, amount=1)
            sheet["A" + str(projectIndex)].value = project
            print("List updated succesfully.")

    return wb.save("spreadSheets/UpdatedbasecampProjects.xlsx")

updateList()