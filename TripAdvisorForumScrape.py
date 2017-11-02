import openpyxl, os
import requests
from lxml import html
import time

'''Fetches all url from the excel file from "SheetName"'''

# workbook = openpyxl.load_workbook(excelPath)
# worksheet = workbook.get_sheet_by_name(SheetName)
filename='NYC _100_Link_To_FreeLancer.xlsx'
excelPath=(os.path.dirname(__file__))+'/'+filename
print(excelPath)
workbook = openpyxl.load_workbook(excelPath)
worksheet = workbook.get_active_sheet()

website = []
for row in worksheet.iter_rows(min_row=1, min_col=3, max_col=3):  ##Getting all the websites,102 max row
    for cell in row:
        website.append(cell.value)

'''End of first block'''


class TripAdvisorXpathForUser1:
    """First Website"""

    def __init__(self, seq_index00):
        self.seqIndex = seq_index00
        self.pageContent = requests.get(website[self.seqIndex - 1])
        self.tree = html.fromstring(self.pageContent.content)
        self.Title = self.tree.xpath('//div[@class="firstPostBox"]//div[@class="postTitle"]/span/text()')
        self.User1 = self.tree.xpath('//div[@class="username"]/a/span/text()')
        self.xPath1 = '//span[@class="postNum" and contains(text(),\"'
        self.user1xpath2 = '.\")]/../../../..//div[@class="username"]/a/span/text()'
        self.date1xPath2 = '.\")]/../../../..//div[@class="postDate"]/text()'  # v02
        self.location1xPath2 = '.\")]/../../../..//div[@class="username"]/..//div[@class="location"]/text()'
        self.posts1xPath2 = '.\")]/../../../..//div[@class="postBadge badge"]/span/text()'
        self.reviewsxPath2 = '.\")]/../../../..//div[@class="reviewerBadge badge"]/span/text()'
        self.replies1xPath2 = '.\")]/../../../..//div[@class="postBody"]/text()'
        self.User1 = self.tree.xpath('//div[@class="username"]/a/span/text()')
        self.Date1 = self.tree.xpath('//div[@class="postDate"]/text()')
        self.Location1 = self.tree.xpath('//div[@class="username"]/..//div[@class="location"]/text()')
        self.Posts1 = self.tree.xpath('//div[@class="postBadge badge"]/span/text()')
        self.Reviews1 = self.tree.xpath('//div[@class="reviewerBadge badge"]/span/text()')
        self.Replies1 = self.tree.xpath('//div[@class="postBody"]')
        self.Rep1 = []
        for i in range(len(self.User1)):
            self.Rep1.append(self.Replies1[i].text_content())

    def firstPost(self):
        firstPostValues = {}
        print("Printing the FirstUser Post")
        firstPostValues = {1: self.Title[0],
                           2: self.User1[0],
                           3: self.Date1[0],
                           4: self.Location1[0],
                           5: self.Posts1[0],
                           6: self.Reviews1[0],
                           7: self.Rep1[0]
                           }
        print("Printing the FirstUser Post")
        for index in range(1,8):
            colnumber = 3+index
            try:
                worksheet.cell(row=self.seqIndex, column=colnumber).value = firstPostValues[index]
            except KeyError as e:
                    print("Unable to print value becuase of %s and at index %d"%(e,index))
        print(self.seqIndex,firstPostValues, sep="|", end="\n")

#######
    def forumReplyPost(self):
        print("Printing Community Post")
        print("Forum Users")
        Users = {}
        for seq in range(1, len(self.Date1)):  # len(Date1))
            # print("seq", seq)
            try:
                if not self.tree.xpath(self.xPath1 + str(seq) + self.user1xpath2):  # == '':
                    # print("if", seq)
                    # print("Invalid xpath with no value")
                    Users[seq] = "No Value"
                else:
                    # print("else", seq)
                    Users[seq] = self.tree.xpath(self.xPath1 + str(seq) + self.user1xpath2)
            except IndexError as e:
                print("IndexError occurred at Last Index seq in Users was", seq)
                Users[seq] = "No Value"
                continue
            except:
                print("Except Seq", seq)
        #print("In the End, Users are", Users)

        print("Forum Dates")
        Dates = {}
        for seq in range(1, len(self.Date1)):  # len(Date1))
            # print("seq", seq)
            try:
                if not self.tree.xpath(self.xPath1 + str(seq) + self.date1xPath2):
                    # print("if", seq)
                    # print("Invalid xpath with no value")
                    print(self.tree.xpath(self.xPath1 + str(seq) + self.date1xPath2))
                    Dates[seq] = "No Value"
                else:
                    # print("else", seq)
                    Dates[seq] = self.tree.xpath(self.xPath1 + str(seq) + self.date1xPath2)
            except IndexError as e:
                print("IndexError occurred at Last Index seq in Dates was", seq)
                Dates[seq] = "No Value"
                continue
            except:
                print("Dates:Except Seq", seq)
        #print("In the End, Dates are", Dates)

        print("Forum Locations")
        location1 = {}
        for seq in range(1, len(self.Date1)):
            try:
                if not self.tree.xpath(self.xPath1 + str(seq) + self.location1xPath2):  # == '':
                    location1[seq] = "No Value"
                else:
                    location1[seq] = self.tree.xpath(self.xPath1 + str(seq) + self.location1xPath2)
            except IndexError as e:
                print("IndexError occurred at Last Index seq in location1 was", seq)
                location1[seq] = "No Value"
                continue
            except:
                print("UnExepected in location", seq)

        print("Forum posts1")
        posts1 = {}
        for seq in range(1, len(self.Date1)):
            try:
                if not self.tree.xpath(self.xPath1 + str(seq) + self.posts1xPath2):  # == '':
                    # print("if", seq)
                    # print("Invalid xpath with no value")
                    posts1[seq] = "No Value"
                else:
                    # print("else", seq)
                    posts1[seq] = self.tree.xpath(self.xPath1 + str(seq) + self.posts1xPath2)
            except IndexError as e:
                print("Last Index seq was", seq)
                posts1[seq] = "No Value"
                continue
            except:
                print("Except Seq", seq)

        print("Forum reviews")
        reviews = {}
        for seq in range(1, len(self.Date1)):  # len(Date1))
            # print("seq", seq)
            try:
                if not self.tree.xpath(self.xPath1 + str(seq) + self.reviewsxPath2):  # == '':
                    # print("if", seq)
                    # print("Invalid xpath with no value")
                    reviews[seq] = "No Value"
                else:
                    # print("else", seq)
                    reviews[seq] = self.tree.xpath(self.xPath1 + str(seq) + self.reviewsxPath2)
            except IndexError as e:
                print("Last Index seq was", seq)
                continue
            except:
                print("Except Seq", seq)
        print("Forum replies1")
        replies1 = {}
        for seq in range(1, len(self.Date1)):
            try:
                if not self.Rep1[seq]:
                    print("if", seq)
                    print("Invalid xpath with no value")
                    replies1[seq] = "No Value"
                else:
                    # print("else", seq)
                    replies1[seq] = self.Rep1[seq]
            except IndexError as e:
                print("IndexError occurred at Last Index seq in replies1 was", seq)
                replies1[seq] = "No Value"
                continue
            except:
                print("Except Seq", seq)
        print("Insertion in excel started")
        for index, colnumber in zip((range(1, len(self.Date1))), range(11, 11 + (len(self.Date1) * 6), 6)):
            try:
                print("formuReplyPost", index, colnumber, self.User1[index], self.Date1[index],
                      self.Location1[index],self.Posts1[index], self.Reviews1[index],
                      self.Rep1[index], sep="|", end="\n")
                worksheet.cell(row=self.seqIndex, column=colnumber + 0).value = str(Users[index])
                worksheet.cell(row=self.seqIndex, column=colnumber + 1).value = str(Dates[index])
                worksheet.cell(row=self.seqIndex, column=colnumber + 2).value = str(location1[index])
                worksheet.cell(row=self.seqIndex, column=colnumber + 3).value = str(posts1[index])
                worksheet.cell(row=self.seqIndex, column=colnumber + 4).value = str(reviews[index])
                worksheet.cell(row=self.seqIndex, column=colnumber + 5).value = str(replies1[index])

            except IndexError:
                print("IndexErrorException unhandled at index %d while insertion"%(index))
                continue
            except KeyError as KE:
                print("Dictionay cannot reference this key at index:%d",index,KE)

for rownumber in range(2, len(website)):
    print("Rownumber value=%d" % (rownumber))
    print("Website", website[rownumber - 1])
    TripAdvisorXpathObj = TripAdvisorXpathForUser1(rownumber)
    TripAdvisorXpathObj.firstPost()
    time.sleep(2)
    TripAdvisorXpathObj.forumReplyPost()
    time.sleep(2)

workbook.save(excelPath)