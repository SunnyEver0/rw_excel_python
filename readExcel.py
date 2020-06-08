import openpyxl
import openpyxl,pprint

wb = openpyxl.load_workbook('censuspopdata.xlsx')

'''
{
 'AL': 
    {
        'Autauga': {'tract':13, 'pop2010':885456}
        'Baldwin': {}
        'Bibb': {}
    }
 'AK': 
    {
 
    }
}
countyData['AL']
'''
# Read the spreadsheet data
sheet = wb.active

countyData = {}

# Fill in countyData with each city's pop and tracts
# range 函数包含start 不包含end 所以要+1
for row in range(2, sheet.max_row+1):
    # Each row in the spreadsheet with data
    state = sheet['B' + str(row)].value
    county = sheet['C' + str(row)].value
    pop = sheet['D' + str(row)].value

    # make sure the key state existed
    # countyData[state] = {}
    countyData.setdefault(state, {})
    # make sure the key for county in state existed
    countyData[state].setdefault(county, {'tract': 0, 'pop': 0})
    # each row represents one census tract, so increment by one
    countyData[state][county]['tract'] += 1
    # Increase the county pop by the pop in census tract
    countyData[state][county]['pop'] += int(pop)

# Open a new text file and write the contents as countyData to it
print('writing results')
# 新建一个文件
resultFile = open('census2010.py', 'w')
# 输入数据
resultFile.write('allData = ' + pprint.pformat(countyData))


