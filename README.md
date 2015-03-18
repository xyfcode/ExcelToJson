# ExcelToJson
开发游戏常用的，Excel中的数据导出到JSON格式的文件中。

# 项目介绍
本工具使用了c#语言开发，使用了第三方库NPOI，是一个在window平台运行的winform程序。主要源代码在Form1.cs文件中。

#如何使用
直接下载压缩包，解压到本地。就可以在window平台运行。
其中EXCEL的格式如下:

|关卡编号|关卡名称|消耗体力|开启等级|开始剧情|
|-------| --------| ------- | ------ | -------|
|GateId|GateName|PowerCost|LvLimite|BeginStoryId|
|String|String|Int|Int|String|
|gate0101|英雄出世|	5|	1|	NewSt002|
|gate0102|初遇黄巾|	5|	2|	NewSt004|
|gate0103|人公将军|	5|	3|	
|gate0104|地公将军|	5|	4|	
|gate0105|天公将军|	5|	5|	
|gate0201|温酒斩华雄|5|	6|	NewSt006|
|gate0202|少帝之死|5 |7|	
|gate0203|曹操的急行军|5 |7|	


程序主要解析第二列和第五列及以下内容。其中第二列是字段标题。第五列及以下是数据。第一列是给策划看的，第三列和第四列是给客户端使用的，可以无视。
导出的格式如下

```
{
   "GATE":{
		"1":{
		  "GateId":"gate0101",
		  "PowerCost":"5",
		  "LvLimited":"1",
		  "War":"war010101",
		  "DropInfo":"3,s1e11",
		  "WarriorExp":"31"
		},
		"2":{
		  "GateId":"gate0102",
		  "PowerCost":"5",
		  "LvLimited":"2",
		  "War":"war010201",
		  "DropInfo":"3,s1e11",
		  "WarriorExp":"32"
		},
		"3":{
		  "GateId":"gate0103",
		  "PowerCost":"5",
		  "LvLimited":"3",
		  "War":"war010301",
		  "DropInfo":"3,s1e61",
		  "WarriorExp":"33"
		},
		"4":{
		  "GateId":"gate0104",
		  "PowerCost":"5",
		  "LvLimited":"4",
		  "War":"war010401",
		  "DropInfo":"3,s1e71",
		  "WarriorExp":"34"
		},
		"5":{
		  "GateId":"gate0105",
		  "PowerCost":"5",
		  "LvLimited":"5",
		  "War":"war010501",
		  "DropInfo":"3,s1e71|3,s1e31",
		  "WarriorExp":"35"
		},
		"6":{
		  "GateId":"gate0201",
		  "PowerCost":"5",
		  "LvLimited":"6",
		  "War":"war020101,war020102,war020103",
		  "DropInfo":"3,s1e11|3,s1e61",
		  "WarriorExp":"36"
		},
		"7":{
		  "GateId":"gate0202",
		  "PowerCost":"5",
		  "LvLimited":"7",
		  "War":"war020201,war020202,war020203",
		  "DropInfo":"3,s1e71|3,s1e81",
		  "WarriorExp":"38"
		},
		"8":{
		  "GateId":"gate0203",
		  "PowerCost":"5",
		  "LvLimited":"7",
		  "War":"war020301,war020302,war020303",
		  "DropInfo":"3,s1e21|3,s1e91",
		  "WarriorExp":"39"
		}
  }
}  
```
