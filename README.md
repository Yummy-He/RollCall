# 随机点名器使用说明（必读）
## 1.表格格式
### 1.1表格位置
学生表格与此RollColl.m文件置于同一文件夹，若在外部，请自行修改启动命令及代码。  
表格名称最好为英文，避免不必要报错。  
### 1.2表格区域
确保学号与姓名列相邻，姓名在学号相邻右侧，且读取区域只包含学号与姓名，不需要包括标题。  
如下表读取区域为B8:C9  
|  | B | C |
| --- | --- | --- |
| 7 | 学号 | 姓名 |
| 8 | 1234 | 张三 |
| 9 | 5678 | 李四 |
### 1.3代码修改
```matlab
% 尝试读取Excel文件，从第8行开始读取，读取第2列(ID)和第3列(Name)
try
    data = readtable(fullfile(pwd, StudentsFile), ...
        'ReadRowNames', false, ...
        'ReadVariableNames', false, ...
        'Range', 'B8:C149');  % 只读取第2列和第3列，从B列到C列
catch ME
    disp('读取Excel文件失败，请检查文件路径或格式是否正确。');
    disp(ME.message);
    return;
end

% 写入点名学生对应行的结果列，从第8行开始
writetable(cell2table(excelData), fullfile(pwd, StudentsFile), ...
    'WriteVariableNames', false, ...
    'Range', [calloutCol, '8:', calloutCol, '149']);  % 写入从第8行开始的指定列
```
## 2.表格格式
