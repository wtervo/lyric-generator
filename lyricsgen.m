%load titles of each word column and all words
[titval, titles, raw] = xlsread('sanat.xlsx', 'A1:P1');
[numval, words, raw2] = xlsread('sanat.xlsx', 'A3:P22');
col_num = numel(words(1, :));
cell_tmp = cell(1, col_num);
%create an empty structure with column titles as fieldnames
word_stru = cell2struct(cell_tmp, titles, 2);

%remove empty cells and add fixes cell arrays into the structure
%under respective fieldname
for m = 1 : col_num
    indx = find(~cellfun(@isempty, words(:, m)));
    new = words(indx, m);
    word_stru.(titles{m}) = new;
end

%open file for writing and get fileID
fid = fopen('biisi.txt', 'w');

%generate parts of the song
titl = titbrid(word_stru, titles, 'titletemp.txt', col_num);
bridge = titbrid(word_stru, titles, 'bridgetemp.txt', col_num);
chorus = versechorus(word_stru, titles, 'chorustemp.txt', col_num, 4);
verse1 = versechorus(word_stru, titles, 'versetemp.txt', col_num, 6, 1);
verse2 = versechorus(word_stru, titles, 'versetemp.txt', col_num, 6, 2);

arr = cell(8, 1);
cat = numel(arr);
[arr{:}] = deal(titl, verse1, chorus, verse2, chorus, bridge, chorus, chorus);

%add every line to the .txt file individually
for m = 1 : cat
    ele = numel(arr{m});
    % special treatment if there is only one element in the cell, like the title
    if ~iscell(arr{m})
        fprintf(fid, '%s\r\n', arr{m});
    else
        for n = 1 : ele
            fprintf(fid, '%s\r\n', arr{m}{n});
        end
    end
    fprintf(fid, '%s\r\n', '');
end

%close .xlsx file
fclose(fid);

%create title or bridge
function titbr = titbrid(stru, titles, file, col_num)

    temp = importdata(file);
    N = numel(temp);
    sent = temp{randi(N)};
    for m = 1 : col_num
        sent = pickword(stru, sent, titles{m});
    end
    sent(1) = upper(sent(1));
    %if bridgetemp file is used, insert the BRIDGE text to make .txt file clearer
    if findstr(file, 'bridge')
        arr = [{'BRIDGE'}; sent];
        titbr = arr;
    else
        titbr = sent;
    end
end

%create verses or chorus
function vercho = versechorus(stru, titles, file, col_num, num, varargin)

    arr = cell(num, 1);
    temp = importdata(file);
    N = numel(temp);
    %nested loops to create num amount of verse lines, which are each unique
    for i = 1 : num
        sent = temp{randi(N)};
        for m = 1 : col_num
            sent = pickword(stru, sent, titles{m});
        end
        %uppercase for the beginning of the sentence
        sent(1) = upper(sent(1));
        arr{i} = sent;
    end
    
    %if varargin has been set, create VERSE text to make the .txt file clearer
    if nargin == 6
        arr = [sprintf('VERSE %u', varargin{1}); arr];
    %else create CHORUS text
    else
        arr = ['CHORUS'; arr];
    end
    vercho = arr;
end

%picks words to the template from the appropriate
%columns that were extracted from the .xlsx file
function sent = pickword(stru, temp, symb)
    
    %replace the appropriate symbols with %s so random words can be inserted
    temp = strrep(temp, symb, '%s');
    t_num = numel(findstr(temp, '%s'));
    prev = 0;
    %get data from the appropriate structure field
    dat = stru.(symb);
    N = length(dat);
    arr = cell(1, t_num);
    for n = 1 : t_num
        num = randi(N);
        %while loop to avoid the same word repeating twice
        while num == prev
            num = randi(N);
        end
        arr{n} = dat{num};
        prev = num;
    end
    %insert random words from arr to the template
    sent = sprintf(temp, arr{:});
    
end