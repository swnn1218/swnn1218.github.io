clc
clear

%读取
M_03=zeros(31,12,46);
for i=1:25
    M=xlsread("D:\【设计院\雨量资料\03上村站降雨资料.xls",num2str(i+74),'C2:N36');
    M=M(4:34,:);%掐头去尾
    M_03(:,:,i)=M;
    [P,I] = max(M,[],'all','linear');
    P_max(i)=P;%年最大降雨量
    mm(i)=ceil(I/31);%发生月
    dd(i)=rem(I,31);%发生日
    if dd(i)==0
        dd(i)=31;
    end
end
for i=26:46
    M=xlsread("D:\【设计院\雨量资料\03上村站降雨资料.xls",num2str(i+1974),'C2:N36');
    M=M(4:34,:);
    M_03(:,:,i)=M;
    [P,I] = max(M,[],'all','linear');
    P_max(i)=P;%年最大降雨量
    mm(i)=ceil(I/31);%发生月
    dd(i)=rem(I,31);%发生日
    if dd(i)==0
        dd(i)=31;
    end
end

%每年转成一列
R=reshape(M_03(:,:,1),[],1);

%% 写入
%% 上村站
%75年
xlswrite("D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx",R);
xlsRenameSheet(['D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx'],1,'75');
%76-20年
for i=2:46
    R=reshape(M_03(:,:,i),[],1);
    xlswrite("D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx",R,num2str(i+74));
end
%给2000-2020sheet重命名
for i=26:46
    xlsRenameSheet(['D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx'],i,num2str(i+1974));
end
myCell = {' ','年最大降水量，(mm)', '月份', '日期'};
yy=[1975:2020];
xlswrite("D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx",myCell,'年最大降雨量','A1');
xlswrite("D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx",yy','年最大降雨量','A2');
xlswrite("D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx",P_max','年最大降雨量','B2');
xlswrite("D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx",mm','年最大降雨量','C2');
xlswrite("D:\【设计院\雨量资料\一列_03上村站降雨资料.xlsx",dd','年最大降雨量','D2');

%% 流域年最大降雨量
myCell2 = {'上村','石龙', '黄牛埔', '松木山', '常平', '同沙', '流域平均'};
xlswrite("D:\【设计院\雨量资料\流域年最大降雨量.xlsx",R,'A2:A373');
xlsRenameSheet(['D:\【设计院\雨量资料\流域年最大降雨量.xlsx'],1,'75');
xlswrite("D:\【设计院\雨量资料\流域年最大降雨量.xlsx",myCell2,'75','A1');

for i=2:46
    R=reshape(M_03(:,:,i),[],1);
    xlswrite("D:\【设计院\雨量资料\流域年最大降雨量.xlsx",R,num2str(i+74),'A2:A373');
    xlswrite("D:\【设计院\雨量资料\流域年最大降雨量.xlsx",myCell2,num2str(i+74),'A1');
end
%给2000-2020sheet重命名
for i=26:46
    xlsRenameSheet(['D:\【设计院\雨量资料\流域年最大降雨量.xlsx'],i,num2str(i+1974));
end
