import os
yourPath = '/Users/nick/Desktop/nick/code/python/ppt/'
allFileList = os.listdir(yourPath)

# 逐一查詢檔案清單

for file in allFileList:

#   這邊也可以視情況，做檔案的操作(複製、讀取...等)
#   使用isdir檢查是否為目錄
#   使用join的方式把路徑與檔案名稱串起來(等同filePath+fileName)
  if os.path.isdir(os.path.join(yourPath,file)):
    print("I'm a directory: " + file)

#   使用isfile判斷是否為檔案
  elif os.path.isfile(yourPath+file):
    print(file)
    
  else:
    print('OH MY GOD !!')