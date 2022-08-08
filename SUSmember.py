#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep 23 11:06:08 2020

@author: xuehuzhou

Update on Wed Mar 16 17:20::15 2022
@author:jibingyu
加入积分等

Update on Wed Aug 8 13:31::48 2022
@author:jibingyu
加入学习次数
"""

import pandas as pd
import sys

studentTYName = "本院团员总名单.xlsx" #团员名单
studentQZName="本院非团员总名单.xlsx" #非团员名单
allStudentFile = "汇总/"+sys.argv[1] +".xlsx" #反馈的excel
Flag=sys.argv[2] #判定是否有效的标志位

All_Student = "总名单.xlsx"
All_Student2 = "学习次数统计/学习次数统计.xlsx"

allStudentFileDf = pd.read_excel(allStudentFile)
studentTYNameDf = pd.read_excel(studentTYName)
studentQZNameDf=pd.read_excel(studentQZName)

All_StudentDf =pd.read_excel(All_Student)
All_StudentDf2 =pd.read_excel(All_Student2,index_col=0)
#All_StudentDf.set(All_StudentDf.dropna())

print(sys.argv[1]+"大学习各支部参与情况如下：")

#这里用于消除反馈excel里面的命名不规范问题，原版有一bug已修改
allClass = allStudentFileDf["姓名"].str.replace(r"[0-9]*"," ")
allClass = allClass.str.replace("-"," ")
allClass = allClass.str.replace("G"," ")
allClass = allClass.str.replace("+"," ")
allClass = allClass.str.replace(" ","")
allClass = set(allClass.dropna())#dropna用于去除含有NaN的行，即反馈excel中有信息缺失则去除此行
className = list(studentTYNameDf.columns)#每一列的首行
#moreInfo = pd.DataFrame(index=className,columns=['总人数','参与人数','参与率','排名','是否合格','本期分数','累计分数',未参加人员'])
moreInfo = pd.DataFrame(index=className,columns=['团员人数','参与人数','参与率','排名','本期分数','累计分数'])
ImpInfo = pd.DataFrame(index=className,columns=['未参与人员'])
all_score = pd.read_excel("支部累计分数/支部累计分数.xlsx",index_col=0)#指定行索引使用第0列


#用于临时存储的变量
allQZCount=[]
allTYCount = []
allCount=[]
attendCount = []
percentCount = []
impresent = []
ispass = []
now_score = []
percentCount1 = []

CT_learner=[]

for i in range(0,len(className)):
    nowClassName = str(className[i])
    nowClass = set(studentTYNameDf[nowClassName].dropna())#i班内所有的团员
    nowQZClass = set(studentQZNameDf[nowClassName].dropna())#非团员

    All_name = list(All_StudentDf[nowClassName].dropna())
    All_name2 = list(All_StudentDf2[nowClassName+"学习次数"])
    Sum_TY = len(nowClass)
    Sum_FTY = len(nowQZClass)
    allCount.append(Sum_TY+Sum_FTY)
    allTYCount.append(Sum_TY)
    allQZCount.append(Sum_FTY)
    #learner=nowClass.intersection(allClass)+nowQZClass.intersection(allClass)#得到i班学习者列表
    learner =list(nowClass.intersection(allClass)|nowQZClass.intersection(allClass))
    #print(nowClassName,learner)


    for i in range(0,len(learner)):
        for j in range(0,len(All_name)):
            if All_name[j]==(learner[i]) and Flag=='updata':
                All_name2[j]=All_name2[j]+1

            if All_name[j]==(learner[i]) and Flag=='back':
                All_name2[j]=All_name2[j]-1

    All_StudentDf2[nowClassName+"学习次数"]=All_name2
    All_StudentDf2.to_excel("学习次数统计/学习次数统计.xlsx")
    #print(All_name2)

    Sum_learn = len(learner)
    attendCount.append(Sum_learn)
    a3 = Sum_learn/Sum_TY
    percentCount.append(round(a3,4))
    percentCount1.append(str(round(a3,2)*100)+"%")
    a4 = list(nowClass.difference(nowClass.intersection(allClass)))
    impresent.append(a4)
    a5 = 0 if a3<0.85 else 1 if a3>=0.85 and a3 <0.9 else 2 if a3 >=0.9 and a3<0.95 else 3 if a3 >=0.95 and a3<1 else 5
    now_score.append(a5)
    #ispass.append("是") if a5>0 else ispass.append("否")
    print('%s共%d名团员%d名非团员，其中%d名同学参与学习，参与率为：%.2f'%(nowClassName,Sum_TY,Sum_FTY,Sum_learn,a3))

#moreInfo['总人数'] = allCount
moreInfo['团员人数'] = allTYCount
#moreInfo['非团员'] = allQZCount
moreInfo['参与人数'] = attendCount
moreInfo['参与率'] = percentCount
#moreInfo['是否合格'] = ispass
moreInfo['本期分数'] = now_score
print(now_score)
moreInfo['累计分数'] = (all_score + pd.DataFrame(now_score,columns=['累计分数'],index=className))['累计分数']
if Flag=='updata':
  all_score['累计分数'] += now_score
  all_score[sys.argv[1]+'分数'] = now_score
if Flag=='back':
  all_score['累计分数'] -= now_score
  all_score[sys.argv[1]+'分数'] = now_score

#moreInfo['未参加人员'] = impresent
ImpInfo['未参与人员'] = impresent
moreInfo.sort_values(by=["参与率",'累计分数'],inplace=True,ascending=[False,False])
moreInfo['排名'] = moreInfo['参与率'].rank(method='first',ascending=False)
moreInfo['参与率'] = moreInfo['参与率'].apply(lambda x: format(x, '.2%'))
moreInfo.to_excel("结果/"+sys.argv[1]+"大学习各支部参与情况.xlsx")
ImpInfo.to_excel("未参与/"+sys.argv[1]+"未参与大学习各支部情况.xlsx")
all_score.to_excel("支部累计分数/支部累计分数.xlsx")
print("EXCELL文件保存成功")




