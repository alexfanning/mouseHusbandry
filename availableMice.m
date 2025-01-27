counter = 2;
for i = params.endOfBreeders+1:size(data{2},1)
    if strcmpi(data{2}{i,"Genotype"},'Hom')
        cellForm = ['A', num2str(counter)];
        avail = {data{2}{i,"Id"},data{2}{i,"CageNum"},data{2}{i,"Age"}{1},data{2}{i,"Sex"}{1}};
        writecell(avail,params.fileName,'Sheet','available','Range',cellForm)
        counter = counter + 1;
    end
end

