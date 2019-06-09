function ceshi_Excel_picture
%设定测试Excel文件名和路径
filespec_user=[pwd '\测试.xls'];
%判断Excel是否已经打开，若已打开，就在打开的Excel中进行操作，
%否则就打开Excel
try
Excel=actxGetRunningServer('Excel.Application');
catch
Excel = actxserver('Excel.Application');
end;
%设置Excel属性为可见
set(Excel, 'Visible', 1);
%返回Excel工作簿句柄
Workbooks = Excel.Workbooks;
%若测试文件存在，打开该测试文件，否则，新建一个工作簿，并保存，文件名为测试.Excel
if exist(filespec_user,'file');
Workbook = invoke(Workbooks,'Open',filespec_user);
else
Workbook = invoke(Workbooks, 'Add');
Workbook.SaveAs(filespec_user);
end
%返回工作表句柄
Sheets = Excel.ActiveWorkBook.Sheets;
%返回第一个表格句柄
sheet1 = get(Sheets, 'Item', 1);
%激活第一个表格
invoke(sheet1, 'Activate');
%如果当前工作表中有图形存在，通过循环将图形全部删除
Shapes=Excel.ActiveSheet.Shapes;
if Shapes.Count~=0;
for i=1:Shapes.Count;
Shapes.Item(1).Delete;
end;
end;

%随机产生标准正态分布随机数，画直方图，并设置图形属性
zft=figure('units','normalized','position',...
[0.280469 0.553385 0.428906 0.251302],'visible','off');
set(gca,'position',[0.1 0.2 0.85 0.75]);
data=normrnd(0,1,1000,1);
hist(data);
grid on;
xlabel('考试成绩');
ylabel('人数');
%将图形复制到粘贴板
hgexport(zft, '-clipboard');
%将图形粘贴到当前表格的A5:B5栏里
Excel.ActiveSheet.Range('A5:B5').Select;
Excel.ActiveSheet.Paste;
%删除图形句柄
delete(zft);