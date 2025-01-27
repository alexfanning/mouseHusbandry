%   Compute mouse colony tasks
%
%   Created by Liv Shalom and Alex Fanning on 1/17/24
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

clear; close all

% Change to folder with excel file
cd('/Users/asf2200/Library/CloudStorage/Dropbox-Alex/Alex Fanning/Kuo lab/Basic Science/Mice breeding/')

params = struct();
params.fileName = fullfile(uigetfile('*.xlsx'));

% Read excel file and convert into data tables
params.sheetNames = sheetnames(params.fileName);
data = cell(1);
for i = 1:length(params.sheetNames) - 2
  data{i} = readtable(params.fileName,'Sheet',params.sheetNames{i});
end
params.numTabs2idx = length(params.sheetNames) - size(data{1},1) - 1;

params.emptyMatrix = cell(100,100);
writecell(params.emptyMatrix,params.fileName,'Sheet','toDoList','Range','A3')

%% Compute animal colony logic

% Create variables for decision tree
params.currWkEnd = datetime('now','Format','MM-dd-yyyy') + days(6);
params.decision = 'Pair';
params.cellFormat = cell(1);
params.rowCounter = [2,2,2,1,2];
params.counter = cell(1);
params.pairNeed = cell(1,10);
params.pairId = [];
params.pair = cell(1,10);

counter = 2;
counter2 = 2;

% Outermost loop iterates through each mouse line
for i = 2:params.numTabs2idx+1
    params.weanIdx = [];
    params.rowIdxs = [];
    params.numBreeders = 0;
    params.groups = cell(1);
    params.groupSizes = [];

    params.rowCounter = params.rowCounter + 1;
    params.cellFormat{9} = ['B', num2str(params.rowCounter(1))];
    params.cellFormat{10} = ['H', num2str(params.rowCounter(2))];
    params.cellFormat{11} = ['N', num2str(params.rowCounter(3))];
    params.cellFormat{12} = ['P', num2str(params.rowCounter(5))];
    writematrix(params.sheetNames{i}, params.fileName,'Sheet','toDoList','Range',params.cellFormat{9});
    writematrix(params.sheetNames{i}, params.fileName,'Sheet','toDoList','Range',params.cellFormat{10});
    writematrix(params.sheetNames{i}, params.fileName,'Sheet','toDoList','Range',params.cellFormat{11});
    writematrix(params.sheetNames{i}, params.fileName,'Sheet','toDoList','Range',params.cellFormat{12});

    params.rowCounter = params.rowCounter + 1;
    params.rowCounter(4) = 1;

    % Find row where the last pair date exists
    params.endOfBreeders = find(data{i}{:,"Id"}==111111,1,'First');

    % Loop through each row in the data
    for row = 1:size(data{i}, 1)
    

        % Calculate 'DateToRemove' based on 'Dob'
        params.dataType = class(data{i}{row,"Dob"});
        if strcmpi('datetime',params.dataType)
            if strcmpi('NaT',cellstr(data{i}{row,"Dob"}))
            elseif strcmpi('NaN',cellstr(data{i}{row,"Dob"}))
            else
                params.cellFormat{1} = ['P' num2str(row+3)];
                if strcmpi(data{i}{row,"Sex"},'F')
                    params.temp = datetime(data{i}{row,"Dob"} + days(data{1,1}{i-1,"FemaleEndOfRotationAge"})*30);
                else
                    params.temp = datetime(data{i}{row,"Dob"} + days(data{1,1}{i-1,"MaleEndOfRotationAge"})*30);
                end
                params.date2remove = datetime(params.temp,'Format','MM/dd/yyyy');
                writecell(cellstr(params.date2remove),params.fileName,'Sheet',params.sheetNames{i},'Range',params.cellFormat{1})
                params.breedersRemoved(params.rowIdxs) = row;
                params.breedersRemSex(params.rowIdxs) = data{i}{row,"Sex"};
                params.rowIdxs = params.rowIdxs + 1;
            end
        end

        % Calculate 'WeanDate' based on 'LitterDob'
        params.dataType = class(data{i}{row,"LitterDob"});
        if strcmpi('datetime',params.dataType)
            if strcmpi('NaT',cellstr(data{i}{row,"LitterDob"}))
            else
                if strcmpi(params.sheetNames{i},'17J') || strcmpi(params.sheetNames{i},'154Q')
                    params.temp = datetime(data{i}{row,"LitterDob"},'Format','MM-dd-yyyy') + days(28);
                elseif strcmpi(params.sheetNames{i},'82Q')
                    params.temp = datetime(data{i}{row,"LitterDob"},'Format','MM-dd-yyyy') + days(19);
                else
                    params.temp = datetime(data{i}{row,"LitterDob"},'Format','MM-dd-yyyy') + days(21);
                end
                
                params.weanDate = datestr(params.temp,'mm/dd/yyyy');
                params.cellFormat{2} = ['L' num2str(row+3)];
                writecell(cellstr(params.weanDate),params.fileName,'Sheet',params.sheetNames{i},'Range',params.cellFormat{2})
            end
        end

        % Calculate age of subjects
        params.dataType = class(data{i}{row,"Dob"});
        if strcmpi('datetime',params.dataType)
            if strcmpi('NaT',cellstr(data{i}{row,"Dob"}))
            else
                params.ageRange = [datetime(data{i}{row,"Dob"},'Format','dd-MMM-yyyy') datetime('now','Format','dd-MMM-yyyy')];
                params.age = caldiff(params.ageRange,{'weeks','days'});
                params.cellFormat{3} = ['E' num2str(row+3)];
                writecell(cellstr(params.age),params.fileName,'Sheet',params.sheetNames{i},'Range',params.cellFormat{3})
            end
        end

        % Automatically calculate monitoring dates for litter births
        params.dataType = class(data{i}{row,"PairingDate"});
        if strcmpi('datetime',params.dataType)
            if strcmpi('NaT',cellstr(data{i}{row,"PairingDate"}))
            else
                % Calculate start of monitoring for potential births
                params.monitor = datetime(data{i}{row,"PairingDate"},'Format','MM/dd/yyyy') + days(19);
                params.cellFormat{4} = ['H' num2str(row+3)];
                writecell(cellstr(params.monitor),params.fileName,'Sheet',params.sheetNames{i},'Range',params.cellFormat{4})
                
                % Calculate end of monitoring for births
                params.endOfMonitor = datetime(params.monitor) + days(19);
                params.cellFormat{5} = ['I' num2str(row+3)];
                writecell(cellstr(params.endOfMonitor),params.fileName,'Sheet',params.sheetNames{i},'Range',params.cellFormat{5})
            end
        end

        % Extract data of female mice with eligible wean dates
        params.dataType = class(data{i}{row,"WeanDate"});
        if strcmpi('datetime',params.dataType) || iscell(data{i}{row,"WeanDate"})
            if datetime(data{i}{row,"WeanDate"},'Format','MM/dd/yyyy') <= params.currWkEnd
    
                % Get the 3-day window around the wean date
                params.temp = datetime(data{i}{row,"WeanDate"}) - days(1);
                params.windowStart = datetime(params.temp,'Format','MM/dd/yyyy');
                params.windowStart = datestr(params.windowStart,'mmm dd');
                params.temp = datetime(data{i}{row,"WeanDate"}) + days(1);
                params.windowEnd = datetime(params.temp,'Format','MM/dd/yyyy');
                params.windowEnd = datestr(params.windowEnd,'mmm dd');
    
                % Calculate the cell placement in Excel to export data
                params.cellFormat{6} = ['A', num2str(params.rowCounter(1))];
        
                % Determine if female should be taken out of breeding rotation
                [params] = endOfRotation(data{i},params,row,i);
    
                % Write weaning info to Excel file
                if strcmpi('82Q_FVB_BL6_RYR1_154Q_colony2.xlsx',params.fileName)
                    params.weanSxs = {data{i}{row,"Id"}, data{i}{row,"CageNum"}{1}, params.windowStart, params.windowEnd, params.decision};
                else
                    params.weanSxs = {data{i}{row,"Id"}, data{i}{row,"CageNum"}, params.windowStart, params.windowEnd, params.decision};
                end
                    writecell(params.weanSxs, params.fileName,'Sheet','toDoList','Range',params.cellFormat{6});
    
                params.rowCounter(1) = params.rowCounter(1) + 1;
                params.weanIdx(params.rowCounter(1)-2) = row;
            end

            % Extract data for mice that should be taken out of rotation
            if row <= params.endOfBreeders
                
                if datetime(data{i}{row,"DateToRemove"},'Format','MM/dd/yyyy') <= params.currWkEnd && ~ismember(row,params.weanIdx)
                    params.cellFormat{7} = ['G', num2str(params.rowCounter(2))];
        
                    params.temp = datetime(data{i}{row,"DateToRemove"}) - days(3);
                    params.windowStart = datetime(params.temp,'Format','MM/dd/yyyy');
                    params.windowStart = datestr(params.windowStart,'mmm dd');
                    params.temp = datetime(data{i}{row,"DateToRemove"}) + days(3);
                    params.windowEnd = datetime(params.temp,'Format','MM/dd/yyyy');
                    params.windowEnd = datestr(params.windowEnd,'mmm dd');
        
                    [params] = endOfRotation(data{i},params,row,i);
                    if strcmpi('82Q_FVB_BL6_RYR1_154Q_colony2.xlsx',params.fileName)
                        params.outOfRotate = {data{i}{row,"Id"},data{i}{row,"CageNum"}{1},params.windowStart,params.windowEnd, params.decision};
                    else
                        params.outOfRotate = {data{i}{row,"Id"},data{i}{row,"CageNum"},params.windowStart,params.windowEnd, params.decision};
                    end
                    writecell(params.outOfRotate,params.fileName,'Sheet','toDoList','Range',params.cellFormat{7})
        
                    params.rowCounter(2) = params.rowCounter(2) + 1;
                end
            end
        end

        % Check for birth cages
        params.currDay = datetime('now','Format','dd-MMM-yyyy');
        params.dataType = class(data{i}{row,"MonitoringDate"});
        if strcmpi('double',params.dataType)
        elseif params.currDay >= datetime(data{i}{row,"MonitoringDate"}{1}) && params.currDay <= datetime(data{i}{row,"EndOfMonitoring"}{1})
            if strcmpi('NaT',cellstr(data{i}{row,"PairingDate"}))
            elseif strcmpi('NaN',cellstr(data{i}{row,"PairingDate"}))
            else

                params.cellFormat{8} = ['M' num2str(params.rowCounter(3))];
                if strcmpi('82Q_FVB_BL6_RYR1_154Q_colony2.xlsx',params.fileName)
                    params.check = {data{i}{row,"Id"},data{i}{row,"CageNum"}{1}};
                else
                    params.check = {data{i}{row,"Id"},data{i}{row,"CageNum"}};
                end
                writecell(params.check,params.fileName,'Sheet','toDoList','Range',params.cellFormat{8})
    
                params.rowCounter(3) = params.rowCounter(3) + 1;
            end
        end

        % Transfer sac'd mice data to old array

    end

    % Create breeding pairs
    % Find relevant mouse lines row of data from 1st sheet in Excel
    params.lineMatch = find(strcmpi(params.sheetNames{i},data{1,1}{:,"Line"}),1,"first");
    try

        % Find the number of breeding pairs based on number of males
        for ii = 1:params.endOfBreeders
            if strcmpi('M',data{i}{ii,"Sex"})
                params.numBreeders = params.numBreeders + 1;
            end
        end

        % Find how many mice are in each breed pairing
        params.groups = find(isnan(data{i}{:,"Id"}));
        params.groupSizes(1) = params.groups(1) - 1;
        for ii = 2:length(params.groups)
            params.groupSizes(ii) = (params.groups(ii) - params.groups(ii-1)) - 1;
        end

        % Find the last breeding group
        params.groupCutoff = find(params.groups > params.endOfBreeders,1,'First') - 1;

        if params.numBreeders > 2
            params.breedSize = median(params.groupSizes(1:params.groupCutoff));
        else
            params.breedSize = max(params.groupSizes(1:params.groupCutoff));
        end

        % Calculate the number of breeding pairs after end of rotation
        % mice have been removed
        if ~isempty(params.breedersRemSex)
            for ii = 1:length(params.breedersRemSex)
                if strcmpi('M',params.breedersRemSex(ii))
                    params.numBreeders = params.numBreeders - 1;
                end
            end
        end

        % Find males/females that need to be re-paired or have 3rd mouse
        for ii = 1:params.groupCutoff
            if params.groupSizes(ii) < params.breedSize
                if ii == 1
                    params.list = 1:params.groups(ii)-1;
                else
                    params.list = params.groups(ii-1)+1:params.groups(ii)-1;
                end
                params.counter{i}{1,ii} = find(strcmpi('M', data{i}{params.list,"Sex"}));
                params.counter{i}{2,ii} = find(strcmpi('F', data{i}{params.list,"Sex"}));
                if sum(params.counter{i}{1,ii}) == 1 && sum(params.counter{i}{2,ii}) == 1
                    params.pairNeed{i}{1,params.rowCounter(4)} = 'female';
                    params.pairNeed{i}{2,params.rowCounter(4)} = params.groups(ii) - params.groupSizes(ii);
                elseif sum(params.counter{i}{1,ii}) >= 1 && sum(params.counter{i}{2,ii}) == 0
                    params.pairNeed{i}{1,params.rowCounter(4)} = 'female';
                    params.pairNeed{i}{2,params.rowCounter(4)} = params.groups(ii) - params.groupSizes(ii);
                elseif sum(params.counter{i}{2,ii}) >= 1 && sum(params.counter{i}{1,ii}) == 0
                    params.pairNeed{i}{1,params.rowCounter(4)} = 'male';
                    params.pairNeed{i}{2,params.rowCounter(4)} = params.groups(ii) - params.groupSizes(ii);
                end
                params.rowCounter(4) = params.rowCounter(4) + 1;
            end
        end

        % Calculate number of breeding pairs needed to match
        % desired number of breeding pairs
        params.numPairs2breed = data{1,1}{params.lineMatch,"NumBreedingPairs"} - params.numBreeders;

        % Pair mice that are already setup as breeders but are w/out mates
        ii = 1;
        if ~isempty(params.pairNeed{i})
            for ii = 1:size(params.pairNeed{i},2)
                j = 1; params.cont = 'n';
                if ~isempty(params.pairNeed{i}{1,ii})
                    params.pair{i}{ii} = {data{i}{params.pairNeed{i}{2,ii},"Id"},data{i}{params.pairNeed{i}{2,ii},"Sex"}{1},data{i}{params.pairNeed{i}{2,ii},"CageNum"},data{i}{params.pairNeed{i}{2,ii},"Genotype"}{1},data{i}{params.pairNeed{i}{2,ii},"Dob"}};
                    if strcmpi(params.pairNeed{i}{1,ii},'male')
                        if ii+1 <= size(params.pairNeed{i},2)
                            for k = ii+1:size(params.pairNeed{i},2)
                                temp = find(strcmpi(data{i}{params.pairNeed{i}{2,k},'Sex'},'M'),1,'first');
                                if ~isempty(temp)
                                    params.pair{i}{ii}(2,:) = {data{i}{params.pairNeed{i}{2,k},"Id"},data{i}{params.pairNeed{i}{2,k},"Sex"}{1},data{i}{params.pairNeed{i}{2,k},"CageNum"},data{i}{params.pairNeed{i}{2,k},"Genotype"}{1},data{i}{params.pairNeed{i}{2,k},"Dob"}};
                                    params.pairNeed{i}{1,k} = [];
                                    params.pairNeed{i}{2,k} = [];
                                    params.numPairs2breed = params.numPairs2breed - 1;
                                else
                                    for t = size(data{i},1):-1:params.endOfBreeders+1
                                        temp1 = strcmpi(data{i}{t,'Sex'},'M');
                                        temp2 = str2double(regexprep(data{i}{t,'Age'}{1},'\D','')) > 8;
                                        %temp3
                                        if temp1 == 1 && temp2 == 1
                                            result{i,ii}(1,t) = t;
                                        end
                                    end
                                    tempId = find(result{i,ii},1,'first');
                                    while params.cont == 'n'
                                        if ismember(tempId,params.pairId) && ~isempty(tempId)
                                            tempId = result{i,ii}(1,1+j);
                                            j = j + 1;
                                        else
                                            params.cont = 'y';
                                        end
                                    end
                                    params.pairId(1,ii) = tempId;
                                    params.pair{i}{ii}(2,:) = {data{i}{params.pairId(1,ii),"Id"},data{i}{params.pairId(1,ii),"Sex"}{1},data{i}{params.pairId(1,ii),"CageNum"},data{i}{params.pairId(1,ii),"Genotype"}{1},data{i}{params.pairId(1,ii),"Dob"}};
                                    params.numPairs2breed = params.numPairs2breed - 1;
                                end
                            end
                        else
                            params.pair{i}{ii} = {data{i}{params.pairNeed{i}{2,ii},"Id"},data{i}{params.pairNeed{i}{2,ii},"Sex"}{1},data{i}{params.pairNeed{i}{2,ii},"CageNum"},data{i}{params.pairNeed{i}{2,ii},"Genotype"}{1},data{i}{params.pairNeed{i}{2,ii},"Dob"}};
                            for t = size(data{i},1):-1:params.endOfBreeders+1
                                temp1 = strcmpi(data{i}{t,'Sex'},'M');
                                temp2 = str2double(regexprep(data{i}{t,'Age'}{1},'\D','')) > 8;
                                if temp1 == 1 && temp2 == 1
                                    result{i,ii}(2,t) = t;
                                end
                            end
                            tempId = find(result{i,ii},1,'first');
                            while params.cont == 'n'
                                if ismember(tempId,params.pairId) && ~isempty(tempId)
                                    tempId = result{i,ii}(2,1+j);
                                    j = j + 1;
                                else
                                    params.cont = 'y';
                                end
                            end
                            params.pairId(1,ii) = tempId;
                            params.pair{i}{ii}(2,:) = {data{i}{params.pairId(1,ii),"Id"},data{i}{params.pairId(1,ii),"Sex"}{1},data{i}{params.pairId(1,ii),"CageNum"},data{i}{params.pairId(1,ii),"Genotype"}{1},data{i}{params.pairId(1,ii),"Dob"}};
                            params.numPairs2breed = params.numPairs2breed - 1;
                        end
                    else
                        if ii+1 <= size(params.pairNeed{i},2)
                            for k = ii+1:size(params.pairNeed{i},2)
                                temp = find(strcmpi(data{i}{params.pairNeed{i}{2,k},'Sex'},'F'),1,'first');
                                if ~isempty(temp)
                                    params.pair{i}{ii}(2,:) = {data{i}{params.pairNeed{i}{2,k},"Id"},data{i}{params.pairNeed{i}{2,k},"Sex"}{1},data{i}{params.pairNeed{i}{2,k},"CageNum"},data{i}{params.pairNeed{i}{2,k},"Genotype"}{1},data{i}{params.pairNeed{i}{2,k},"Dob"}};
                                    params.pairNeed{i}{1,k} = [];
                                    params.pairNeed{i}{2,k} = [];
                                    params.numPairs2breed = params.numPairs2breed - 1;
                                else
                                    for t = size(data{i},1):-1:params.endOfBreeders+1
                                        temp1 = strcmpi(data{i}{t,'Sex'},'F');
                                        temp2 = str2double(regexprep(data{i}{t,'Age'}{1},'\D','')) > 6;
                                        %temp3
                                        if temp1 == 1 && temp2 == 1
                                            result{i,ii}(1,t) = t;
                                        end
                                    end
                                    tempId = find(result{i,ii},1,'first');
                                    while params.cont == 'n'
                                        if ismember(tempId,params.pairId) && ~isempty(tempId)
                                            tempId = result{i,ii}(1,1+j);
                                            j = j + 1;
                                        else
                                            params.cont = 'y';
                                        end
                                    end
                                    params.pairId(1,ii) = tempId;
                                    params.pair{i}{ii}(2,:) = {data{i}{params.pairId(1,ii),"Id"},data{i}{params.pairId(1,ii),"Sex"}{1},data{i}{params.pairId(1,ii),"CageNum"},data{i}{params.pairId(1,ii),"Genotype"}{1},data{i}{params.pairId(1,ii),"Dob"}};
                                    params.numPairs2breed = params.numPairs2breed - 1;
                                end
                            end
                        else
                            params.pair{i}{ii} = {data{i}{params.pairNeed{i}{2,ii},"Id"},data{i}{params.pairNeed{i}{2,ii},"Sex"}{1},data{i}{params.pairNeed{i}{2,ii},"CageNum"},data{i}{params.pairNeed{i}{2,ii},"Genotype"}{1},data{i}{params.pairNeed{i}{2,ii},"Dob"}};
                            for t = size(data{i},1):-1:params.endOfBreeders+1
                                temp1 = strcmpi(data{i}{t,'Sex'},'F');
                                temp2 = str2double(regexprep(data{i}{t,'Age'}{1},'\D','')) > 6;
                                if temp1 == 1 && temp2 == 1
                                    result{i,ii}(1,t) = t;
                                end
                            end
                            tempId = find(result{i,ii},1,'first');
                            while params.cont == 'n'
                                if ismember(tempId,params.pairId) && ~isempty(tempId)
                                    tempId = result{i,ii}(1,1+j);
                                    j = j + 1;
                                else
                                    params.cont = 'y';
                                end
                            end
                            params.pairId(1,ii) = tempId;
                            params.pair{i}{ii}(2,:) = {data{i}{params.pairId(1,ii),"Id"},data{i}{params.pairId(1,ii),"Sex"}{1},data{i}{params.pairId(1,ii),"CageNum"},data{i}{params.pairId(1,ii),"Genotype"}{1},data{i}{params.pairId(1,ii),"Dob"}};
                            params.numPairs2breed = params.numPairs2breed - 1;
                        end
                    end
                end
            end
        end
        % Determine if number of current breeding pairs equals or
        % exceeds number of desired breeding pairs
        if params.numBreeders > data{1,1}{params.lineMatch,"NumBreedingPairs"}
            disp('Too many breeding pairs')
        end
            %%%%%%% HAVE TO DO LIST BE NOTIFIED OF TOO MANY BREEDING CAGES
        if params.numPairs2breed > 0
            for m = 1:params.numPairs2breed
                j = 1; params.cont = 'n';
                for t = size(data{i},1):-1:params.endOfBreeders+1
                    params.ageNum = regexp(data{i}{t,'Age'}{1},'\d*','match');
                    temp1 = strcmpi(data{i}{t,'Sex'},'F');
                    temp2 = str2double(params.ageNum{1}) > 6;
                    temp3 = str2double(params.ageNum{1}) < 14;
                    if strcmpi(params.sheetNames{i},'17J')
                        temp4 = strcmpi(data{i}{t,"Genotype"}{1},'Hom');
                    elseif strcmpi(params.sheetNames{i},'154Q')
                        params.check = {data{i}{row,"Id"},data{i}{row,"CageNum"}};
                    end
                    if temp1 == 1 && temp2 == 1 && temp3 == 1
                        result{i,ii+m}(1,t) = t;
                    end
                end
                tempId = find(result{i,ii+m}(1,:),1,'last');
                while params.cont == 'n'
                    if ismember(tempId,params.pairId) && ~isempty(tempId)
                        tempId = result{i,ii+m}(1,1+j);
                        j = j + 1;
                    else
                        params.cont = 'y';
                    end
                end
                params.pairId(1,ii+m) = tempId;
                params.pair{i}{ii+m}(1,:) = {data{i}{params.pairId(1,ii+m),"Id"},data{i}{params.pairId(1,ii+m),"Sex"}{1},data{i}{params.pairId(1,ii+m),"CageNum"},data{i}{params.pairId(1,ii+m),"Genotype"}{1},data{i}{params.pairId(1,ii+m),"Dob"}};
                
                j = 1; params.cont = 'n';
                for t = size(data{i},1):-1:params.endOfBreeders+1
                    params.ageNum = regexp(data{i}{t,'Age'}{1},'\d*','match');
                    temp1 = strcmpi(data{i}{t,'Sex'},'M');
                    temp2 = str2double(params.ageNum{1}) > 8;
                    temp3 = str2double(params.ageNum{1}) < 16;
                    if temp1 == 1 && temp2 == 1 && temp3 == 1
                        result{i,ii+m}(2,t) = t;
                    end
                end
                tempId = find(result{i,ii+m}(2,:),1,'first');
                while params.cont == 'n'
                    if ismember(tempId,params.pairId) && ~isempty(tempId)
                        tempId = result{i,ii+m}(2,1+j);
                        j = j + 1;
                    else
                        params.cont = 'y';
                    end
                end
                params.pairId(2,ii+m) = tempId;
                params.pair{i}{ii+m}(2,:) = {data{i}{params.pairId(2,ii+m),"Id"},data{i}{params.pairId(2,ii+m),"Sex"}{1},data{i}{params.pairId(2,ii+m),"CageNum"},data{i}{params.pairId(2,ii+m),"Genotype"}{1},data{i}{params.pairId(2,ii+m),"Dob"}};
                params.numPairs2breed = params.numPairs2breed - 1;
            end
        end
    catch ME
        strcmp(ME.identifier,'Incorrect number or types of inputs or outputs for function')
    end

    % Write data to 'toDolist'
    if ~isempty(params.pair{i})
        for ii = 1:size(params.pair{i},2)
            if ~isempty(params.pair{i}{ii})
                params.cellFormat{12} = ['P', num2str(params.rowCounter(5))];
                writecell(params.pair{i}{ii},params.fileName,'Sheet','toDoList','Range',params.cellFormat{12})
                params.rowCounter(5) = params.rowCounter(5) + 2;
            end
        end
    end

    if i == 2
        for iii = params.endOfBreeders+1:size(data{2},1)
            if strcmpi(data{2}{iii,"Genotype"},'Hom')
                cellForm = ['A', num2str(counter)];
                avail = {data{2}{iii,"Id"},data{2}{iii,"CageNum"},data{2}{iii,"Age"}{1},data{2}{iii,"Sex"}{1}};
                writecell(avail,params.fileName,'Sheet','available','Range',cellForm)
                counter = counter + 1;
            else
                cellForm2 = ['G', num2str(counter2)];
                avail2 = {data{2}{iii,"Id"},data{2}{iii,"CageNum"},data{2}{iii,"Age"}{1},data{2}{iii,"Sex"}{1}};
                writecell(avail2,params.fileName,'Sheet','available','Range',cellForm2)
                counter2 = counter2 + 1;
            end
        end
    end

end

for pj = 1:size(data{6},1)
    if strcmpi(data{6}{pj,"Genotype"},'Hom')
        cellForm = ['A', num2str(counter)];
        avail = {data{6}{pj,"Id"},data{6}{pj,"CageNum"},data{6}{pj,"Dob"},data{6}{pj,"Sex"}{1}};
        writecell(avail,params.fileName,'Sheet','available','Range',cellForm)
        counter = counter + 1;
    else
        cellForm2 = ['G', num2str(counter2)];
        avail2 = {data{6}{pj,"Id"},data{6}{pj,"CageNum"},data{6}{pj,"Dob"},data{6}{pj,"Sex"}{1}};
        writecell(avail2,params.fileName,'Sheet','available','Range',cellForm2)
        counter2 = counter2 + 1;
    end
end
