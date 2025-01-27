%   Compute mouse colony tasks
%
%   Created by Liv Shalom and Alex Fanning on 1/17/24
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

clear; close all

% Change to folder with excel file
cd('/Users/asf2200/Library/CloudStorage/Dropbox/Kuo lab/Basic Science/Mice breeding')

params = struct();
params.fileName = fullfile(uigetfile('*.xlsx'));

% Read excel file and convert into data tables
params.sheetNames = sheetnames(params.fileName);
data = cell(1);
for i = 1:length(params.sheetNames) - 2
  data{i} = readtable(params.fileName,'Sheet',params.sheetNames{i});
end

params.emptyMatrix = cell(40,40);
writecell(params.emptyMatrix,params.fileName,'Sheet','toDoList','Range','A3')

%% Compute animal colony logic

% Create variables for decision tree
params.currWkEnd = datetime('now','Format','MM-dd-yyyy') + days(6);
params.decision = 'Pair';
params.cellFormat = cell(1);
params.rowCounter = [2,2,2];

% Outermost loop iterates through each mouse line
for i = 1:length(params.sheetNames) - 3
    params.weanIdx = [];

    params.rowCounter = params.rowCounter + 1;
    params.cellFormat{9} = ['B', num2str(params.rowCounter(1))];
    params.cellFormat{10} = ['H', num2str(params.rowCounter(2))];
    params.cellFormat{11} = ['N', num2str(params.rowCounter(3))];
    writematrix(params.sheetNames{i}, params.fileName,'Sheet','toDoList','Range',params.cellFormat{9});
    writematrix(params.sheetNames{i}, params.fileName,'Sheet','toDoList','Range',params.cellFormat{10});
    writematrix(params.sheetNames{i}, params.fileName,'Sheet','toDoList','Range',params.cellFormat{11});

    params.rowCounter = params.rowCounter + 1;

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
                    params.temp = datetime(data{i}{row,"Dob"} + days(180));
                else
                    params.temp = datetime(data{i}{row,"Dob"} + days(240));
                end
                params.date2remove = datetime(params.temp,'Format','MM/dd/yyyy');
                writecell(cellstr(params.date2remove),params.fileName,'Sheet',params.sheetNames{i},'Range',params.cellFormat{1})
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
                params.weanSxs = {data{i}{row,"Id"}, data{i}{row,"CageNum"}, params.windowStart, params.windowEnd, params.decision};
                writecell(params.weanSxs, params.fileName,'Sheet','toDoList','Range',params.cellFormat{6});
    
                params.rowCounter(1) = params.rowCounter(1) + 1;
                params.weanIdx(params.rowCounter(1)-2) = row;
            end

            % Extract data for mice that should be taken out of rotation
            if datetime(data{i}{row,"DateToRemove"},'Format','MM/dd/yyyy') <= params.currWkEnd && ~ismember(row,params.weanIdx)
                params.cellFormat{7} = ['G', num2str(params.rowCounter(2))];
    
                params.temp = datetime(data{i}{row,"DateToRemove"}) - days(3);
                params.windowStart = datetime(params.temp,'Format','MM/dd/yyyy');
                params.windowStart = datestr(params.windowStart,'mmm dd');
                params.temp = datetime(data{i}{row,"DateToRemove"}) + days(3);
                params.windowEnd = datetime(params.temp,'Format','MM/dd/yyyy');
                params.windowEnd = datestr(params.windowEnd,'mmm dd');
    
                [params] = endOfRotation(data{i},params,row,i);
                params.outOfRotate = {data{i}{row,"Id"},data{i}{row,"CageNum"},params.windowStart,params.windowEnd, params.decision};
                writecell(params.outOfRotate,params.fileName,'Sheet','toDoList','Range',params.cellFormat{7})
    
                params.rowCounter(2) = params.rowCounter(2) + 1;
            end
        end

        % Check for birth cages
        params.currDay = datetime('now','Format','MM-dd-yyyy');
        params.dataType = class(data{i}{row,"MonitoringDate"});
        if strcmpi('double',params.dataType)
        elseif params.currDay >= data{i}{row,"MonitoringDate"} && params.currWkEnd <= data{i}{row,"EndOfMonitoring"}
            if strcmpi('NaT',cellstr(data{i}{row,"LitterDob"}))
            elseif strcmpi('NaN',cellstr(data{i}{row,"LitterDob"}))
            else

                params.cellFormat{8} = ['M' num2str(params.rowCounter(3))];
                params.check = {data{i}{row,"Id"},data{i}{row,"CageNum"}};
                writecell(params.check,params.fileName,'Sheet','toDoList','Range',params.cellFormat{8})
    
                params.rowCounter(3) = params.rowCounter(3) + 1;
            end
        end

        % Transfer sac'd mice data to old array

    end

end
