import openpyxl
from openpyxl import Workbook
from google.colab import files

# Task information hardcoded as per the input provided
task_data = [
    {
        "Person": "Mahad",
        "Total Tasks": 1,
        "Completed Tasks": 1,
        "In Progress": 1,
        "In Review": 0,
        "In Revision": 0,
        "To Do": 0,
        "Extra Tasks": 0,
        "Update": "Find a solution to the problem of data capturing and extraction for business leads and other features, via programming paradigm or no code method.",
        "Extra": "None"
    },
    {
        "Person": "Samia",
        "Total Tasks": 1,
        "Completed Tasks": 0,
        "In Progress": 1,
        "In Review": 0,
        "In Revision": 0,
        "To Do": 0,
        "Extra Tasks": 5,
        "Update": "Revamping TriVA's website content - Wrote content for 3 pages (Home, About Us, Blog & Services) and compiled a list of 200 keywords +  20 competitors assessing their Domain Authority and Spam scores.",
        "Extra": "Completed an 11-hour SEO crash course on on-page, off-page, technical SEO, Google Analytics, and Google Search Console. Researched 15+ competitors for a Christmas blog, wrote 1 blog (including meta title, meta description, and images), and drafted 2 samples. Planned and scheduled 12 emails, wrote content for 1, and created an email scheduling document by weeks and dates. Researched and wrote descriptions for 8 GBP Posts. Learned to draft blogs on WordPress."
    },
    {
        "Person": "Asad",
        "Total Tasks": 2,
        "Completed Tasks": 0,
        "In Progress": 1,
        "In Review": 1,
        "In Revision": 0,
        "To Do": 1,
        "Extra Tasks": 0,
        "Update": "Comic strip |ScienceJumma (Design 4 Character Designs and design 4 different comics sample of Comic strip), Research on comic strip and explore different Ai Comic maker Platforms",
        "Extra": ""
    },
    {
        "Person": "Asad",
        "Total Tasks": 2,
        "Completed Tasks": 0,
        "In Progress": 1,
        "In Review": 1,
        "In Revision": 0,
        "To Do": 1,
        "Extra Tasks": 0,
        "Update": "Comic strip |ScienceJumma (Design 4 Character Designs and design 4 different comics sample of Comic strip), Research on comic strip and explore different Ai Comic maker Platforms",
        "Extra": ""
    },
    {
        "Person": "Meher",
        "Total Tasks": 4,
        "Completed Tasks": 1,
        "In Progress": 2,
        "In Review": 1,
        "In Revision": 0,
        "To Do": 0,
        "Extra Tasks": "Learning moosend and how to do birthday automation",
        "Update": "  wrote the titles for the Christmas graphics."
        "I uploaded a story and posted images on Instagram and Facebook."
        "I contacted Moosend via live chat, but they requested I email them regarding the email list query."
        "I completed the business list on Moosend."
        "I finished writing slogans and CTAs for private party hire and uploaded them on Figma."
        "I completed the task of creating slogans and CTAs for the Christmas graphics."
        "I am currently working on researching business directories and updating the sheet, with only the Brighton directories left to add."
        "I completed the research on Brighton directories and updated the sheet for the Christmas task."
        "I am working on setting up Moosend birthday email automation."
        "I am searching for reels, stock videos, and music for the Christmas events ",
        "Extra": ""
    },
    {
        "Person": "Azfar",
        "Total Tasks": 8,
        "Completed Tasks": 8,
        "In Progress": 0,
        "In Review": 0,
        "In Revision": 0,
        "To Do": 0,
        "Extra Tasks": 2,
        "Update": "Responded to reviews"
                  "Posted Christmas Party GBP post"
                  "Drafted 2 blogs: Top 10 Corporate Christmas Gifts" and "Importance of Choosing the Right Christmas Party Venues on WordPress"
                  "Planned email schedule"
                  "Conducted SEO research"
                  "Drafted blog: Top 10 Traditional South Indian Cuisines You Need to Try on WordPress"
                  "Written 1 new email for the missing week"
                  "Writing FAQs for 2 blogs"
                  "Researched landing pages"
                  "Written FAQs and drafted 2 blogs: 10 Unforgettable Vegan Restaurants Near Brighton You Can’t Miss!" and "Discover the Top 5 Must-Try Bengali Cuisine"
                  "Finalised all emails according to the week number",
        "Extra": "Learned about landing pages content and local SEO"
    },
    {
        "Person": "Kamran",
        "Total Tasks": 7,
        "Completed Tasks": 0,
        "In Progress": 1,
        "In Review": 6,
        "In Revision": 0,
        "To Do": 0,
        "Extra Tasks": 1,
        "Update": "Designed social media posts for all the campaigns"
                   "Designed E-mail banner gifs and images for E-mail marketing"
                   "Designed 12 E-mails for 3 campaigns"
                   "Finding stock videos photos and music for Zari’s Instagram" 
                   "Took English class" 
                   "Meeting with 360’s and also with pod member’s",
        "Extra": "Designed 2 social media posts for 360’s Instagram"
    },

]

# Create an Excel workbook and sheet
workbook = Workbook()
sheet = workbook.active
sheet.title = "Team Tasks"

# Write headers to the first row
headers = ["Person", "Total Tasks", "Completed Tasks", "In Progress", "In Review", "In Revision", "To Do", "Extra Tasks", "Update", "Extra"]
sheet.append(headers)

# Write task data to the sheet
for entry in task_data:
    row = [
        entry.get("Person", ""),
        entry.get("Total Tasks", ""),
        entry.get("Completed Tasks", ""),
        entry.get("In Progress", ""),
        entry.get("In Review", ""),
        entry.get("In Revision", ""),
        entry.get("To Do", ""),
        entry.get("Extra Tasks", ""),
        entry.get("Update", ""),
        entry.get("Extra", "")
    ]
    sheet.append(row)

# Save the Excel file
excel_file_path = "team_tasks.xlsx"
workbook.save(excel_file_path)

# Download the file to the local machine
files.download(excel_file_path)

print(f"Data successfully written to {excel_file_path}")
