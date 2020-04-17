| Sheet1 | A     | B          | C         | D    | E          | F       | G     |
| ------ | ----- | ---------- | --------- | ---- | ---------- | ------- | ----- |
| 1      | id    | first name | last name | age  | phone      | record  | score |
| 2      | 12301 | Andy       | Zhang     | 21   | 131***1111 | 3组     | 70    |
| 3      | 12302 | Bob        | Wang      | 22   | 133***1112 | 10次    | 55    |
| 4      | 12303 | Cindy      | Zhao      | 23   | 135***1113 | 完成5组 | 91    |
| 5      | 12304 | David      | Qian      | 24   | 137***1114 | 20次    | 60    |
| 6      | 12305 | David      | Sun       | 25   | 139***1115 | 完成4   | 86    |
| 7      | 12306 | Eason      | Zhang     | 26   | 131***1116 | 5组     | 95    |



计算区域内不同单元格个数=SUMPRODUCT(1/COUNTIF(G2:M2,G2:M2))

模糊匹配=VLOOKUP(""&B1&"",A1:A7,1,0)

分段匹配=LOOKUP(A1, {0,50,60,70,80,90}, {"F","E","D","C","B","A"})

按组计算中位数（https://www.extendoffice.com/zh-CN/documents/excel/4815-excel-pivot-table-median.html）=MEDIAN(IF($B$2:$B$31=B2,$C$2:$C$31))

计算单元格内特定符号的数量=LEN(A1)-LEN(SUBSTITUTE(A1,"*",""))

删除单元格内的隐藏双引号和空格=""&(VALUE(CLEAN(A1)))

提取字符出中的数值https://jingyan.baidu.com/article/ca00d56c3e38a2e99eebcfb7.html
=MIDB(A2,SEARCHB("?",A2),2\*LEN(A2)-LENB(A2))**