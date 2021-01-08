import re
import pandas as pd
import argparse

#讀入相關的參數
ap = argparse.ArgumentParser()
ap.add_argument("-t", "--txt", required=True, help="Path to the txt of the TXT file.")
args = vars(ap.parse_args())

# 開啟檔案
def open_txt_file(filepath):
    f = open(filepath)
    return f

# 找到ballname
def findBallName(ballName, string):
    result = re.findall(f"(UCPU1\.{ballName}\,|UCPU1\.{ballName}$)", string, re.I)
    return result

#step1: 開啟檔案
filepath = args["txt"]
try:
    f = open_txt_file(filepath)
    print(f"Step1: Open file. (File name: {filepath}) -DONE")
except:
    print(f"File is not exit, check file name.")
rawData = []
for i in f:
    row = i.replace("\n","")#去掉換行符號
    rawData.append(row)

#step2: 整理TXT
#找出個$NETS並清除不需要的資料
keyIDX = []
for i in rawData:
    if re.findall(f"\$(NETS|END)", i , re.I):
#         print(re.findall(f"\$.*", i, re.I))
        keyIDX.append(rawData.index(i))
rawData2 = rawData[keyIDX[0] +1: keyIDX[1]]

# 把TXT整理成dictionary
netName, netData = "", ""
netInfo = {}
clearData = {}
# 處理txt檔 分好net name 跟 GPIO資料
for line in rawData2:
#     print(re.findall(f".+\;", line))
    if re.findall(f".+\;", line):
        netName = re.findall(f".+\;", line)[0].replace(";", "")#分號前是net name
        
        #處理 net data 
        netData = re.findall(f"\;.+", line)[0].replace(";", "")#分號後是net data
        netData = netData.split()
        netData = ",".join(netData)


        netInfo = { netName : netData }
        clearData.update(netInfo)
    else:
        line = ",".join(line.split())#用空格白切割 再用,分割
        
        if clearData.__contains__(netName):
            clearData[netName] += line
        else:
            print("erroe", line)

print ("Step2: Sort TXT DATA - DONE.")



# 叫出intel GPIO template 
filepath2 = "GPIO mapping.xlsx"
data = pd.read_excel(filepath2)
CPUBallName = data.iloc[:,0].tolist()

data["netName"] = ""#產生net name行

for netName, netData in clearData.items():
#     print(netData)
    for ballName in CPUBallName:
        result = findBallName(ballName, netData)
#         print(result)
        if result:
            if data.loc[data[data["CPU Ball Name"]==f"{ballName}"].index.tolist()[0], "netName" ] == "":
#                 print("yes")
                data.loc[data[data["CPU Ball Name"]==f"{ballName}"].index.tolist()[0], "netName" ] = netName
            else:            
                print("重複")
                data.loc[data[data["CPU Ball Name"]==f"{ballName}"].index.tolist()[0], "netName" ] += netName
print ("Step3: Mapping - DONE.")

elsxPath = f'{filepath.replace(".txt",".xlsx")}'
write = pd.ExcelWriter(elsxPath)

# save_excel(write, dfFunction, "dfFunction")
data.to_excel(write, sheet_name=f'data', index=False)

write.save()

print ("Step4: Save Result - DONE.")


