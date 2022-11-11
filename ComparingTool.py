#fileListから情報取得
def sprit_fileList(data_Path):
  import os
  #Body
  buf = os.path.basename(data_Path)
  fileName = buf

  countN = buf.find("_")
  Method = buf[ : countN]
  buf = buf[countN+1 : ]

  countN = buf.find("_")
  Material = buf[ : countN]
  buf = buf[countN+1 : ]

  countN = buf.find("_")
  Weight = buf[ : countN]
  buf = buf[countN+1 : ]

  countN = buf.find("_")
  SlidingSpeed = buf[ : countN]
  buf = buf[countN+1 : ]

  countN = buf.find(".")
  number = int(buf[ : countN])

  return fileName, Method, Material, Weight, SlidingSpeed, number

from doctest import OutputChecker
from operator import index
#DB→Excel出力
def export_DB(dataPath, DB, header, save_Path):
  #Excel操作系ライブラリーのインポート
  import pandas
  df = pandas.DataFrame(DB)
  df.to_excel(save_Path + os.sep + dataPath, index=False, header=header)

  frag = "True"
  return frag


def function1(fileList, method_name, delta, save_Path):    #delta=平均する微分の範囲
  import pandas
  import os

  for count_fL in range(len(fileList)):     #count_fL = count_fileList
    print(str(count_fL+1) + "/" + str(len(fileList)))

    #DB取得
    data_Path = fileList[count_fL]
    DB = pandas.read_excel(data_Path, sheet_name = "Sheet1")

    #DB→計算用一次元配列に置き換え
    SlidingTime = []
    FrictionCoefficient = []
    for i in range(len(DB)-19):
      SlidingTime.append(DB.values[i+19][0])
      FrictionCoefficient.append(DB.values[i+19][2])
    
    #fileList→Param取得
    fileName, Method, Material, Weight, SlidingSpeed, number = sprit_fileList(data_Path)

    #DB→differential
    slope = []
    slope_Outcome = []
    for n in range(len(SlidingTime)-1):
      slope.append((FrictionCoefficient[n + 1] - FrictionCoefficient[n]) / (SlidingTime[n + 1] - SlidingTime[n]))
      if n%delta == delta - 1:
        for i in range(delta):
          ave_slope = (slope[n-(delta-1) + i] / delta)
        slope_Outcome.append([SlidingTime[n-(delta-1)], ave_slope])
        ave_slope = 0

    #DB→Excel出力
    new_fileName = method_name + os.sep + fileName
    header = ["time","slope"]
    frag_process = export_DB(new_fileName, slope_Outcome, header, save_Path)
    if frag_process == "True":
      print("seved file.")
    else:
      print("error02\n出力失敗しました。")



def function2(fileList, method_name, delta, save_Path):      #delta=平均する微分の範囲
  import pandas
  
  #Body
  Range_Outcome = []
  for count_fL in range(len(fileList)):     #count_fL = count_fileList
    print(str(count_fL+1) + "/" + str(len(fileList)))

    #DB取得
    data_Path = fileList[count_fL]
    DB = pandas.read_excel(data_Path, sheet_name = "Sheet1")

    #ここにExcelの入力を作成＆countloopを利用
    #fileList→Param取得
    fileName, Method, Material, Weight, SlidingSpeed, number = sprit_fileList(data_Path)

    #DB→計算用一次元配列に置き換え
    SlidingTime = []
    FrictionCoefficient = []
    for i in range(len(DB)-19):
      SlidingTime.append(DB.values[i+19][0])
      FrictionCoefficient.append(DB.values[i+19][2])

    #DB→Range[delta]→average_Range[delta]
    max = 0
    min = 1000
    max_range = 0
    average_range = 0
    for n in range(len(SlidingTime)):
      if FrictionCoefficient[n] > max:
        max = FrictionCoefficient[n]
      if FrictionCoefficient[n] < min:
        min = FrictionCoefficient[n]
      #1区間の計測が終了した場合
      if n%delta == delta-1:
        if (max-min) > max_range:
          max_range = (max-min)
        average_range = average_range + (max - min) *delta / len(SlidingTime)
        max = 0
        min = 1000
    Range_Outcome.append([Method, Material, Weight, SlidingSpeed, number, max_range, average_range])

  #DB→Excel出力
  new_fileName = method_name + ".xlsx"
  header = ["Method", "Material", "Weight", "SlidingSpeed", "number", "max_Range", "average_Range"]
  frag_process = export_DB(new_fileName, Range_Outcome, header, save_Path)
  if frag_process == "True":
    print("seved file.")
  else:
    print("error02\n出力失敗しました。")
  
from numpy import sqrt
def function3(fileList, method_name, delta, save_Path):      #delta=平均する微分の範囲
  import pandas


  #Body
  SD_Outcome = []
  for count_fL in range(len(fileList)):     #count_fL = count_fileList
    print(str(count_fL+1) + "/" + str(len(fileList)))

    #DB取得
    data_Path = fileList[count_fL]
    DB = pandas.read_excel(data_Path, sheet_name = "Sheet1")

    #ここにExcelの入力を作成＆countloopを利用
    #fileList→Param取得
    fileName, Method, Material, Weight, SlidingSpeed, number = sprit_fileList(data_Path)

    #DB→計算用一次元配列に置き換え
    SlidingTime = []
    FrictionCoefficient = []
    for i in range(len(DB)-19):
      SlidingTime.append(DB.values[i+19][0])
      FrictionCoefficient.append(DB.values[i+19][2])

    #DB→Standard Deviation[delta]→average_Standard Deviation[delta]
    #区間平均をとる
    average_FC = 0
    average_SD = 0
    for n in range(len(SlidingTime)):
      average_FC = FrictionCoefficient[n] / delta
      #1区間の計測が終了した場合
      if n%delta == delta-1:
        for i in range(n-19, n):
          #分散の計算
          Distributed = (FrictionCoefficient[i] - average_FC) * (FrictionCoefficient[i] - average_FC) / delta
        #標準偏差の平均
        average_SD = sqrt(Distributed) *delta / len(SlidingTime)
        average_FC = 0

    SD_Outcome.append([Method, Material, Weight, SlidingSpeed, number, average_SD])

  #DB→Excel出力
  new_fileName = method_name + ".xlsx"
  header = ["Method", "Material", "Weight", "SlidingSpeed", "number", "average_Standard Deviation"]
  frag_process = export_DB(new_fileName, SD_Outcome, header, save_Path)
  if frag_process == "True":
    print("seved file.")
  else:
    print("error02\n出力失敗しました。")

def function4(fileList, method_name, save_Path):
  import pandas
  import numpy
  import os


  #Body
  for count_fL in range(len(fileList)):     #count_fL = count_fileList
    StableTime_Outcome = []
    print(str(count_fL+1) + "/" + str(len(fileList)))

    #DB取得
    data_Path = fileList[count_fL]
    DB = pandas.read_excel(data_Path, sheet_name = "Sheet1")

    #ここにExcelの入力を作成＆countloopを利用
    #fileList→Param取得
    fileName, Method, Material, Weight, SlidingSpeed, number = sprit_fileList(data_Path)

    #DB→計算用一次元配列に置き換え
    SlidingTime = []
    FrictionCoefficient = []
    for i in range(len(DB)-19):
      SlidingTime.append(DB.values[i+19][0])
      FrictionCoefficient.append(DB.values[i+19][2])
    if len(SlidingTime) < len(FrictionCoefficient):
      finalRow = len(SlidingTime)
    else:
      finalRow = len(FrictionCoefficient)
      
    #DB→Estimate Time(available for loop)→Estimate Linear Function→相関係数によるFB(相関係数　=0.95)→determine Time
    for EST in range(finalRow):     #EST:EstimateTime
      copySlidingTime = []
      copyFrictionCoefficient = []
      for i in range(finalRow-EST):
        copySlidingTime.append(SlidingTime[i+EST])
        copyFrictionCoefficient.append(FrictionCoefficient[i+EST])
      coef=numpy.polyfit(copySlidingTime, copyFrictionCoefficient, 1)      #determini coefficient, coef[0]=傾き,coef[1]=切片
      copyFrictionCoefficient1 = numpy.poly1d(coef)(copySlidingTime)     
                      #Estimate Linear Function (copySlidingTime, copyFrictionCoefficient1), y=numpy.poly1d(coefficient)(x)
      # リストをps.Seriesに変換
      s1=pandas.Series(copyFrictionCoefficient)
      s2=pandas.Series(copyFrictionCoefficient1)

      # pandasを使用してPearson's rを計算
      res=s1.corr(s2)   # numpy.float64 に格納される

      StableTime_Outcome.append([EST/10, res])

    #DB→Excel出力
    new_fileName = fileName
    header = ["EST", "res"]
    frag_process = export_DB(new_fileName, StableTime_Outcome, header, save_Path)
    if frag_process == "True":
      print("seved file.")
    else:
      print("error02\n出力失敗しました。")

def function5(fileList, method_name, CorrelationCoefficient, save_Path):
  import pandas
  import numpy

  StableTime_Outcome = []
  for count_fL in range(len(fileList)):     #count_fL = count_fileList
    print(str(count_fL+1) + "/" + str(len(fileList)))

    #DB取得
    data_Path = fileList[count_fL]
    DF = pandas.read_excel(data_Path, sheet_name = "Sheet1")

    #ここにExcelの入力を作成＆countloopを利用
    #fileList→Param取得
    fileName, Method, Material, Weight, SlidingSpeed, number = sprit_fileList(data_Path)

    #DB→計算用一次元配列に置き換え
    SlidingTime = []
    FrictionCoefficient = []
    for i in range(len(DF)-19):
      SlidingTime.append(DF.values[i+19][0])
      FrictionCoefficient.append(DF.values[i+19][2])
    finalRow = len(SlidingTime)

    #DB→Estimate Time(available for loop)→Estimate Linear Function→相関係数によるFB(相関係数　=0.95)→determine Time
    for EST in range(0, finalRow):     #EST:EstimateTime
      tempSlidingTime = []
      tempFrictionCoefficient = []
      for i in range(EST, finalRow):
        tempSlidingTime.append(SlidingTime[i])
        tempFrictionCoefficient.append(FrictionCoefficient[i])
      coef=numpy.polyfit(tempSlidingTime, tempFrictionCoefficient, 2)      #determini coefficient, coef[0]=x^2係数,coef[1]=x^1係数, coef[2]=x^0係数
      approximation_FrictionCoefficient = numpy.poly1d(coef)(tempSlidingTime)
              #Estimate Linear Function (copySlidingTime, copyFrictionCoefficient1), y=numpy.poly1d(coefficient)(x)
      # リストをps.Seriesに変換
      s1=pandas.Series(tempFrictionCoefficient)
      s2=pandas.Series(approximation_FrictionCoefficient)
      # pandasを使用してPearson's rを計算
      res=s1.corr(s2)   # numpy.float64 に格納される
      if res > CorrelationCoefficient:
        break
      if EST == finalRow-1:
        EST = "NA"
      print("completion: " + str(count_fL+1) + "/" + str(len(fileList)) + " Calculating.Process("\
         + str(EST/10) + "/" + str(finalRow/10) + ") EST=" + str(EST/10) + " CorrelationCoefficient=" + str(res))
    if not EST == "NA":
      EST = EST /10
    StableTime_Outcome.append([Method, Material, Weight, SlidingSpeed, number, EST, coef[0], coef[1], coef[2]])

    #DB→Excel出力
    new_fileName = method_name + ".xlsx"
    header = ["Method", "Material", "Weight", "SlidingSpeed", "number", "StableTime", "x^2係数", "x^1係数", "x^0係数"]
    frag_process = export_DB(new_fileName, StableTime_Outcome, header, save_Path)
    if frag_process == "True":
      print("seved [" + fileName + "]")
    else:
      print("error02\n出力失敗しました。")





def function6(fileList, method_name, save_Path):
  import pandas
  import numpy

  StableTime_Outcome = []
  for count_fL in range(len(fileList)):     #count_fL = count_fileList
    print(str(count_fL+1) + "/" + str(len(fileList)))

    #DB取得
    data_Path = fileList[count_fL]
    DB = pandas.read_excel(data_Path, sheet_name = "Sheet1")

    #ここにExcelの入力を作成＆countloopを利用
    #fileList→Param取得
    fileName, Method, Material, Weight, SlidingSpeed, number = sprit_fileList(data_Path)

    #DB→計算用一次元配列に置き換え
    SlidingTime = []
    FrictionCoefficient = []
    for i in range(len(DB)-19):
      SlidingTime.append(DB.values[i+19][0])
      FrictionCoefficient.append(DB.values[i+19][2])
    
    finalRow = len(SlidingTime)

    #DB→Estimate Time(available for loop)→Estimate Linear Function→相関係数によるFB(相関係数　=0.95)→determine Time
    output_EST = 0
    output_slope = 0
    output_yIntercept = 0
    output_correlationCoefficien = 0
    for EST in range(0, int(finalRow*2/3)):     #EST:EstimateTime
      copySlidingTime = []
      copyFrictionCoefficient = []
      for i in range(0, finalRow-EST,10):
        copySlidingTime.append(SlidingTime[i+EST])
        copyFrictionCoefficient.append(FrictionCoefficient[i+EST])
      coef=numpy.polyfit(copySlidingTime, copyFrictionCoefficient, 1)      #determini coefficient, coef[0]=傾き,coef[1]=切片
      copyFrictionCoefficient1 = numpy.poly1d(coef)(copySlidingTime)     
              #Estimate Linear Function (copySlidingTime, copyFrictionCoefficient1), y=numpy.poly1d(coefficient)(x)
      # リストをps.Seriesに変換
      s1=pandas.Series(copyFrictionCoefficient)
      s2=pandas.Series(copyFrictionCoefficient1)
      # pandasを使用してPearson's rを計算
      res=s1.corr(s2)   # numpy.float64 に格納される
      if res > output_correlationCoefficien:
        output_EST = EST
        output_slope = coef[0]
        output_yIntercept = coef[1]
        output_correlationCoefficien = res
      print("completion: " + str(count_fL+1) + "/" + str(len(fileList)) + " Calculating.Process("\
         + str(EST/10) + "/" + str(finalRow/10) + ") EST=" + str(EST/10) + " CorrelationCoefficient=" + str(res))
    output_EST = output_EST /10
    StableTime_Outcome.append([Method, Material, Weight, SlidingSpeed, number, output_correlationCoefficien,output_EST, output_slope, output_yIntercept])

    #DB→Excel出力
    new_fileName = method_name + ".xlsx"
    header = ["Method", "Material", "Weight", "SlidingSpeed", "number", "CorrelationCoefficien","StableTime", "slope", "y-intercept"]
    frag_process = export_DB(new_fileName, StableTime_Outcome, header, save_Path)
    if frag_process == "True":
      print("seved [" + fileName + "]")
    else:
      print("error02\n出力失敗しました。")
    

def function7(fileList=None, method_name=None, save_Path=None, number_coef=None):
  import pandas
  import numpy
  for count_fL in range(len(fileList)):     #count_fL = count_fileList
    print(str(count_fL+1) + "/" + str(len(fileList)))
    dataPath = fileList[count_fL]
    DF = pandas.read_excel(dataPath, sheet_name = "Sheet1")
    fileName, Method, Material, Weight, SlidingSpeed, number = sprit_fileList(dataPath)

    #DF→List translate
    SlidingTime = []
    FrictionCoefficient = []
    for i in range(len(DF)-19):
      SlidingTime.append(DF.values[i+19][0])
      FrictionCoefficient.append(DF.values[i+19][2])

    differential = pandas.DataFrame(data=numpy.array([SlidingTime, FrictionCoefficient]), columns=['SlidingTime', 'FrictionCoefficient'])
    temp_differential_value = []
    temp_differential_value.append(0)
    for i in range(1, len(SlidingTime)):
      temp_differential_value.append((FrictionCoefficient[i]-FrictionCoefficient[i-1])/(SlidingTime[i]-SlidingTime[i-1]))
    differential['1_deriv'] = temp_differential_value
    for n in range(2, number_coef+1):
      temp_differential_value = []
      temp_differential_value.append(0)
      for i in range(1, len(SlidingTime)):
        temp_differential_value.append((differential.values[i][n-1]-differential.values[i-1][n-1])/(SlidingTime[i]-SlidingTime[i-1]))
      differential[str(n)+'_deriv'] = temp_differential_value
    differential.to_excel(save_Path + os.sep + os.path.splitext(os.path.basename(dataPath))[0] + '.xlsx', \
      index=False, header=list(differential.columns))
    print("seved [" + fileName + "]")






#摩擦係数の評価
#必要ライブラリー
import sys
import pandas
import glob
import os
import tkinter as tk
import tkinter.filedialog
from tkinter import messagebox

method_index = ["20秒間の平均傾き全表示","20秒間の平均レンジ","20秒間の平均標準偏差","xx秒間の平均レンジ",\
                "xx秒間の平均標準偏差","相関係数全表示_摩擦係数近似","摩擦係数近似による安定時間演算(相関係数0.80)", \
                "摩擦係数近似による安定時間演算(相関係数0.xx)", "摩擦係数近似による安定時間演算(最大相関係数)", "摩擦係数微分(微分回数xx)"]

#処理種別決定&処理内容確認
print("計算手法を入力してください。")
for i in range(len(method_index)):
  print(str(i) + ":" + method_index[i])
frag_method = int(input(">>"))
method_name = method_index[frag_method]

#fileList取得
target_Path = tkinter.filedialog.askdirectory(title="データの参照をフォルダを指定してください。", mustexist=True) + os.sep + "*.xlsx"
fileList = glob.glob(target_Path)
if len(fileList) == 0:
  print("error03\nデータがありません。システムを終了します。")
  sys.exit()

#保存先決定
save_Path = tkinter.filedialog.askdirectory(title="データの保存先を指定してください。", mustexist=True)

#20秒間の平均傾き全表示
if frag_method == 0:
  delta = 200
  function1(fileList, method_name, delta, save_Path)

#20秒間の平均レンジ
elif frag_method ==1:
  delta = 200
  function2(fileList, method_name, delta, save_Path)

#20秒間の平均標準偏差
elif frag_method == 2:
  delta = 200
  function3(fileList, method_name, delta, save_Path)

#xx秒間の平均レンジ
elif frag_method == 3:
  rangeTime = int(input("時間範囲を入力してください。[sec]\n>>"))
  method_name = str(rangeTime) + method_name[2: ]
  delta = rangeTime *10
  function2(fileList, method_name, delta, save_Path)

#xx秒間の平均標準偏差
elif frag_method == 4:
  rangeTime = int(input("時間範囲を入力してください。[sec]\n>>"))
  method_name = str(rangeTime) + method_name[2: ]
  delta = rangeTime *10
  function3(fileList, method_name, delta, save_Path)

#相関係数全表示_摩擦係数近似
elif frag_method == 5:
  function4(fileList, method_name, save_Path)

#摩擦係数近似による安定時間演算(相関係数0.80)
elif frag_method == 6:
  CorrelationCoefficient = 0.80
  function5(fileList, method_name, CorrelationCoefficient, save_Path)

#摩擦係数近似による安定時間演算(相関係数0.xx)
elif frag_method == 7:
  CorrelationCoefficient = float(input("相関係数を入力してください。\n>>"))
  method_name = method_name[ : 20] + str(float(CorrelationCoefficient)) + ')'
  function5(fileList, method_name, CorrelationCoefficient, save_Path)

#摩擦係数近似による安定時間演算(相関係数0.xx)
elif frag_method == 8:
  function6(fileList, method_name, save_Path)

#摩擦係数微分・二回微分
elif frag_method == 9:
  number_coef = int(input('微分回数を入力してください。\n>>'))
  method_name = method_name[:-3] + str(number_coef) + ')'
  function7(fileList=fileList, method_name=method_name, save_Path=save_Path, number_coef=number_coef)

else:
  print("error04\nモードがありません。システムを終了します。")
  sys.exit()

print("All Process is Done\nstop Running\nSee you my Boss")


#ver1.0.0　基本設計作成
#ver1.0.1　変数名変更&ESTコード見直し
#ver1.1.0　微分機能追加
#ver1.1.1　微小な修正
#ver1.1.2　FrictionCoefficient追加
#ver1.1.3　微小な修正