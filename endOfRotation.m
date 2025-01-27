%   Decision tree for mice at the end of breeding rotation
%
%   Written by Alex Fanning on 1/18/24
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function [parameters] = endOfRotation(dataIn,parameters,rowNum,count)

% Determine if mouse should be taken out of breeding rotation
if datetime(dataIn{rowNum,"DateToRemove"},'Format','MM-dd-yyyy') <= parameters.currWkEnd

    if strcmpi('17J',parameters.sheetNames{count}) || strcmpi('RYR1',parameters.sheetNames{count})
    % Decision tree for whether to keep or sacrifice the mouse
    %%%%%%%%% ADD DOUBLE TRANSGENIC TREE %%%%%%%%%%%%%%%%%%
        if strcmpi(dataIn{rowNum,"Genotype"},'Hom')
            parameters.decision = 'Move to 317A';
        elseif strcmpi(dataIn{rowNum,3}, 'Het')
            parameters.decision = 'Sac';
        end
    elseif strcmpi(parameters.sheetNames{count},'82QxPCP2_cre')
        if strcmpi(dataIn{rowNum,"Genotype"},'82QCre+')
            parameters.decision = 'Move to 317A';
        elseif strcmpi(dataIn{rowNum,"Genotype"}, '82QCre-')
            parameters.decision = 'Sac';
        else
            parameters.decision = 'Sac';
        end
    else
        if strcmpi(dataIn{rowNum,"Genotype"},'Het')
            parameters.decision = 'Move to 317A';
        elseif strcmpi(dataIn{rowNum,"Genotype"}, 'Wt')
            parameters.decision = 'Sac';
        end
    end

else
    parameters.decision = 'Pair';
end
