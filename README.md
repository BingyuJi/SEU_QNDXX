# 东南大学自动化学院青年大学习统计

###
代码依托于东南大学无锡校区青年大学习统计方式，网址：https://github.com/huzai9527
在其功能上加入团员非团员统计，分数累计计算，学习次数累计等多种新功能
###

###
加updata后缀写入累计积分、学习次数文件
加back后缀退回累计积分、学习次数文件
其他后缀则会生成未参与名单，但并不会修改积分、学习次数
###

## 环境安装

- 本项目在**python 3** 下运行,并且依赖**pands**,你可以在终端运行以下命令进行安装
```
pip install pandas
```

## 所需要的文件以及格式
- 本院非团员总名单,首行为班级名称,首列为索引号(excel自带),将班级同学名字替换即可
  文件放在和SUSmember.py同级
- 本院团员总名单,首行为班级名称,首列为索引号(excel自带),将班级同学名字替换即可
  文件放在和SUSmember.py同级
- 本院总名单,首行为班级名称,首列为索引号(excel自带),将班级同学名字替换即可
  文件放在和SUSmember.py同级
- 支部累计分数:首行为列名第一列列名空,第二列列名累计分数
  第一列为填班级名,第二列为保存的班级名
  文件需要手动创建放在文件夹支部累计分数
- 学习次数:本院总名单后每列多一个统计列
  名字在原基础上加“学习次数”，初始默认手动设置为0
  文件需要手动创建放在文件夹学习次数统计
- 最后是大学习的统计情况,此文件是上级下发的
  只要移到汇总文件夹目录即可
- **注:** 总名单、班级累计分数、学习次数可以在下载好的文件中直接修改

## 安装好环境以及准备好相应文件后,在终端运行(文件目录路径下)
```
python SUSmember.py 自动化学院-202205 updata
```
- 运行完成输出，各班详细情况，并在结果文件夹生成EXCELL统计文件


## 有不是团员的只要在团员名单中删除，非团员名单加入
