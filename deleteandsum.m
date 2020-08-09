clc,clear
warning off
% 根据names.txt文件中的名字，删除当前目录下不在里面的名字的文件，然后提取出
% 每一个表格内某一个信息，最后在将信息总结至一个输出excel文件

%%     读取 names.txt中的名字
fid = fopen('names.txt');
Names = {};
name_num = 1;
while(1)
    line_ex = fgetl(fid);
    if(line_ex == -1)
        break;
    end
    name = [line_ex '.xls'];
    Names{name_num} = name;
    name_num = name_num + 1;
end
fclose(fid);

%%  删除文件夹中的".xls"不在"names.txt"中的文件
xls_name=dir(fullfile('*.xls'));
current_files_count = length(xls_name);
for i = 1:current_files_count
    now_name = xls_name(i).name;
    if(ismember(now_name,Names))
        outputs = ['----',now_name,'*****************留下'];
        disp(outputs);
    else
        outputs = ['----',now_name,'-----------------删除'];
        disp(outputs);
        delete(now_name)
    end
end

%% 对当前文件夹下的所有文件进行信息提取
xls_name=dir(fullfile('*.xls'));
current_files_count = length(xls_name);
% select_keynames = {'股东名称','股东及出资信息'};
select_keynames = {'姓名'};
company_name = {};
company_name_count = 1;

for i = 1:current_files_count
    now_name = xls_name(i).name;
    outputs = ['信息提取 ',now_name,' 中'];
    %disp(outputs);
    A=readtable(now_name,'Format','auto', 'PreserveVariableNames',1 ,'ReadRowNames',true);
    var_count = length(A.Properties.VariableNames);
    names_list = [];
    names_list_count = 1;
    
    for j = 1:var_count
        lists = eval(['A.Var' num2str(j)]);
        try
            flag = ismember(lists,select_keynames);
            if(any(flag))
                first_po = find(flag,1);
                for k = (first_po+1):length(lists)
                    if(isempty(char(lists(k))))
                        break
                    else
                        names_list = [names_list ' ' char(lists(k))];
                        names_list_count = names_list_count + 1;
                    end
                end
            end
            
        catch
            
        end
        company_name(i,:) = {now_name(1:(end-4)),names_list};
    end
end
T = table(company_name);
writetable(T,'test.xls')
