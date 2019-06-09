function ceshi_Excel_picture
%�趨����Excel�ļ�����·��
filespec_user=[pwd '\����.xls'];
%�ж�Excel�Ƿ��Ѿ��򿪣����Ѵ򿪣����ڴ򿪵�Excel�н��в�����
%����ʹ�Excel
try
Excel=actxGetRunningServer('Excel.Application');
catch
Excel = actxserver('Excel.Application');
end;
%����Excel����Ϊ�ɼ�
set(Excel, 'Visible', 1);
%����Excel���������
Workbooks = Excel.Workbooks;
%�������ļ����ڣ��򿪸ò����ļ��������½�һ���������������棬�ļ���Ϊ����.Excel
if exist(filespec_user,'file');
Workbook = invoke(Workbooks,'Open',filespec_user);
else
Workbook = invoke(Workbooks, 'Add');
Workbook.SaveAs(filespec_user);
end
%���ع�������
Sheets = Excel.ActiveWorkBook.Sheets;
%���ص�һ�������
sheet1 = get(Sheets, 'Item', 1);
%�����һ�����
invoke(sheet1, 'Activate');
%�����ǰ����������ͼ�δ��ڣ�ͨ��ѭ����ͼ��ȫ��ɾ��
Shapes=Excel.ActiveSheet.Shapes;
if Shapes.Count~=0;
for i=1:Shapes.Count;
Shapes.Item(1).Delete;
end;
end;

%���������׼��̬�ֲ����������ֱ��ͼ��������ͼ������
zft=figure('units','normalized','position',...
[0.280469 0.553385 0.428906 0.251302],'visible','off');
set(gca,'position',[0.1 0.2 0.85 0.75]);
data=normrnd(0,1,1000,1);
hist(data);
grid on;
xlabel('���Գɼ�');
ylabel('����');
%��ͼ�θ��Ƶ�ճ����
hgexport(zft, '-clipboard');
%��ͼ��ճ������ǰ����A5:B5����
Excel.ActiveSheet.Range('A5:B5').Select;
Excel.ActiveSheet.Paste;
%ɾ��ͼ�ξ��
delete(zft);