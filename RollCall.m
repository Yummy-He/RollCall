function RollCall(StudentsFile)

% 读取Excel文件，从第8行开始读取，读取第2列(ID)和第3列(Name)
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

% 手动为table设置列名，但不会写入Excel中
data.Properties.VariableNames = {'ID', 'Name'};

% 计算学生总人数
numStudents = height(data);

% 询问点名多少人
numToCall = input(['当前学生总数为 ', num2str(numStudents), ' 人，请输入要点名的人数（最少1人，最多', num2str(numStudents), '人）: ']);

% 检查输入是否有效
if ~isnumeric(numToCall) || numToCall < 1 || numToCall > numStudents
    disp('点名人数输入无效，请输入1到学生总人数之间的数字。');
    return;
end

% 询问是否显示学号
showID = input('是否保护隐私不显示学号？(1: 显示, 0: 不显示): ', 's');
if length(showID) > 1
    disp('输入错误，只应输入一个数字。');
    return;
end
showID = str2double(showID);

% 询问是第几次点名
calloutNum = input('这是第几次点名？: ');
if ~isnumeric(calloutNum) || calloutNum < 1
    disp('点名次数输入错误，请输入一个大于0的整数。');
    return;
end

% 生成随机点名索引，按用户指定人数点名
randomIndices = randperm(numStudents, numToCall);

% 选择被点名学生的信息
selectedStudents = data(randomIndices, {'ID', 'Name'});

% 显示点名信息
disp('点名学生信息:');
for i = 1:numToCall
    if showID
        fprintf('%d: 学号: %s, 姓名: %s\n', i, selectedStudents.ID{i}, selectedStudents.Name{i});
    else
        fprintf('%d: 姓名: %s\n', i, selectedStudents.Name{i});
    end
end

% 确认到否
responses = input(['请输入每位学生的到场情况（如输入长度为', num2str(numToCall), '的字符串，1表示到，0表示不到）: '], 's');
if length(responses) ~= numToCall || ~all(ismember(responses, ['1' '0']))
    disp('输入反馈错误，请输入长度正确且只包含1和0的字符串。');
    return;
end

% 准备回写到Excel中（"1" 表示到，不填；"0" 表示不到，写"缺"）
excelData = cell(numStudents, 1);  % 创建一个空单元格数组，用于写入Excel
for i = 1:numToCall
    if responses(i) == '0'
        excelData{randomIndices(i)} = '缺';  % 如果没到，写"缺"
    end
end

% 将点名结果回写到Excel表格中

% 确定点名结果要写入的列（从J列开始）
startCol = 'J';
calloutCol = char(startCol + (calloutNum - 1));
% 写入点名学生对应行的结果列，从第8行开始
writetable(cell2table(excelData), fullfile(pwd, StudentsFile), ...
    'WriteVariableNames', false, ...
    'Range', [calloutCol, '8:', calloutCol, '149']);  % 写入从第8行开始的指定列

% 显示完成信息
disp('点名结果已写回Excel文件。');
end