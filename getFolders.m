%   Get relevant folders
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function [listOutput,numFolders] = getFolders(type)

pat = lettersPattern;
listOutput = dir;
numFolders = 0;

if type == 0
    for i = 1:length(listOutput)

        if startsWith(listOutput(i).name,'.')
            continue
        elseif endsWith(listOutput(i).name, '.mat')
            continue
        elseif endsWith(listOutput(i).name,'.fig')
            continue
        else
            numFolders = numFolders + 1;
            totalList(numFolders) = i;
        end

    end
elseif type == 1
    for i = 1:length(listOutput)
        
        if startsWith(listOutput(i).name,'.')
            continue
        else
            numFolders = numFolders + 1;
            totalList(numFolders) = i;
        end
        
    end
elseif type == 2
    for i = 1:length(listOutput)
        if startsWith(listOutput(i).name,'.')
            continue
        elseif endsWith(listOutput(i).name, '.mat')
            continue
        elseif startsWith(listOutput(i).name,'~')
            continue
        else
            numFolders = numFolders + 1;
            totalList(numFolders) = i;
        end
    end

elseif type == 3
    for i = 1:length(listOutput)
        if startsWith(listOutput(i).name,'.')
            continue
        elseif endsWith(listOutput(i).name, '.mat')
            continue
        elseif startsWith(listOutput(i).name,'~')
            continue
        elseif contains(listOutput(i).name,'.xlsx')
            continue
        elseif startsWith(listOutput(i).name,',')
            continue
        else
            numFolders = numFolders + 1;
            totalList(numFolders) = i;
        end
    end
elseif type == 4
    for i = 1:length(listOutput)
        if startsWith(listOutput(i).name,'~')
            continue
        elseif contains(listOutput(i).name,'.xlsx')
            numFolders = numFolders + 1;
            totalList(numFolders) = i;
        end
    end
elseif type == 5
    for i = 1:length(listOutput)
        if startsWith(listOutput(i).name,'~')
            continue
        elseif startsWith(listOutput(i).name,'.')
            continue
        elseif endsWith(listOutput(i).name,'.fig')
            continue
        elseif contains(listOutput(i).name,'.xlsx')
            continue
        elseif contains(listOutput(i).name,'compiled')
            continue
        elseif endsWith(listOutput(i).name, '.mat')
            numFolders = numFolders + 1;
            totalList(numFolders) = i;
        end
    end
end

if exist('totalList')
    listOutput = listOutput(totalList);
else
    listOutput = listOutput;
end
