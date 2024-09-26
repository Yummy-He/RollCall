# 随机点名器使用说明（必读）
## 1.表格格式
### 1.1表格位置
学生表格与此 RollCall.m 文件置于 Matlab 工作文件夹，若在外部，请自行修改启动命令及代码。  
表格名称最好为英文（如 “Students.xlsx”），避免不必要报错。  
### 1.2表格区域
确保学号与姓名列相邻，姓名在学号相邻右侧，且读取区域只包含学号与姓名，不需要包括标题。  
如下表，读取区域为 B8:C9，点名结果将保存在 JK...列（可修改）中：
  |  | B | C | ... | J | K |
  | :---: | :---: | :---: | :---: | :---: | :---: |
  | 7 | 学号 | 姓名 |  | 第1次 | 第2次 |
  | 8 | 1234 | 张三 |  |  |  |
  | 9 | 5678 | 李四 |  |  |  |
  | ... | ... | ... |  |  |  |
  | 123 | 5679 | 刘六 |  |  |  |
  | ... | ... | ... |  |  |  |
  | 149 | 1011 | 王五 |  |  |  |
### 1.3代码修改
在 RollCall.m 中根据实际情况，阅读注释，**在如下部位，修改表格读写区域**：
```matlab
3.   % 尝试读取Excel文件，从第8行开始读取，读取第2列(ID)和第3列(Name)
4.   try
5.       data = readtable(fullfile(pwd, StudentsFile), ...
6.           'ReadRowNames', false, ...
7.           'ReadVariableNames', false, ...
8.           'Range', 'B8:C149');  % 只读取第2列和第3列，从B列到C列
9.   catch ME
10.      disp('读取Excel文件失败，请检查文件路径或格式是否正确。');
11.      disp(ME.message);
12.      return;
13.  end

78.  % 确定点名结果要写入的列（从J列开始）
79.  startCol = 'J';
80.  calloutCol = char(startCol + (calloutNum - 1));

81.  % 写入点名学生对应行的结果列，从第8行开始
82.  writetable(cell2table(excelData), fullfile(pwd, StudentsFile), ...
83.      'WriteVariableNames', false, ...
84.      'Range', [calloutCol, '8:', calloutCol, '149']);  % 写入从第8行开始的指定列
```
## 2.使用说明
### 2.1准备
打开Matlab，选择工作文件夹，将 RollCall.m 和表格（如 “Students.xlsx”）**放在工作文件夹**。  
### 2.1使用
在**命令行**窗口中输入，根据**文件名称**自行修改启动命令：
```matlab
RollCall('Students.xlsx');
```
回车。  


此时自动计算总人数，并要求输入点名人数，如
```matlab
当前学生总数为 142 人，请输入要点名的人数（最少1人，最多142人）:
```
输入后，回车。  
  
  
询问是否显示学号，如
```matlab
当前学生总数为 142 人，请输入要点名的人数（最少1人，最多142人）: 5
是否保护隐私不显示学号？(1: 显示, 0: 不显示): 
```
输入后，回车。
  
  
询问是第几次点名，在本示例中，填1则结果保存在 J 列，如
```matlab
当前学生总数为 142 人，请输入要点名的人数（最少1人，最多142人）: 5
是否保护隐私不显示学号？(1: 显示, 0: 不显示): 1
这是第几次点名？: 
```
输入后，回车。
  
  
展示被点到的学生，如：
```matlab
当前学生总数为 142 人，请输入要点名的人数（最少1人，最多142人）: 5
是否保护隐私不显示学号？(1: 显示, 0: 不显示): 1
这是第几次点名？: 1
点名学生信息:
1: 学号: *******, 姓名: ***
2: 学号: *******, 姓名: ***
3: 学号: *******, 姓名: ***
4: 学号: *******, 姓名: ***
5: 学号: *******, 姓名: ***
请输入每位学生的到场情况（如输入长度为5的字符串，1表示到，0表示不到）:
```
此时教师依次叫名字，**到输入1，缺勤输入0**，多少人就输入多少位，请勿多输入或少输入。  
输入后，回车。
  
  
本例中点名5人，若都缺勤，则输入00000，如：
```matlab
当前学生总数为 142 人，请输入要点名的人数（最少1人，最多142人）: 5
是否保护隐私不显示学号？(1: 显示, 0: 不显示): 1
这是第几次点名？: 1
点名学生信息:
1: 学号: *******, 姓名: ***
2: 学号: *******, 姓名: ***
3: 学号: *******, 姓名: ***
4: 学号: *******, 姓名: ***
5: 学号: *******, 姓名: ***
请输入每位学生的到场情况（如输入长度为5的字符串，1表示到，0表示不到）: 00000
点名结果已写回 Excel 文件。
```
  
  
此时点名完成，结果回写入表格。  
本例中，第一次点名结果储存在 J 列，到则不显示，缺勤则填入“缺”。如：  
  |  | B | C | ... | J | K |
  | :---: | :---: | :---: | :---: | :---: | :---: |
  | 7 | 学号 | 姓名 |  | 第1次 | 第2次 |
  | 8 | 1234 | 张三 |  | 缺 |  |
  | 9 | 5678 | 李四 |  |  |  |
  | ... | ... | ... |  |  |  |
  | 123 | 5679 | 刘六 |  | 缺 |  |
  | ... | ... | ... |  |  |  |
  | 149 | 1011 | 王五 |  |  |  |
