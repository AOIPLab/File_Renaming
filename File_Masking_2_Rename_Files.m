%% Replace file names according to a provided spreadsheet of old and new 
% file names
% Not approved for AOIP use. Last validated: Never
% Updated - 2019.04.17 - Alex Salmon
% Updated - 2020.10.23 - Brian Higgins: Added counter for # of missing
% files and end of file summary report
%   Use recursive file search

%% Imports
addpath('lib');


%% Constants
OLD_LABEL = 'Old';
NEW_LABEL = 'New';
xl_fname = 'RenameCode.xlsm';


%% Get root directory and do recursive file search
root_dir = uigetdir;
srch_results = subdir(root_dir);
srch_ffnames = {srch_results.name}';
[all_paths, all_names, all_ext] = cellfun(@fileparts, srch_ffnames, ...
    'uniformoutput', false);
all_fnames = cellfun(@(x, y) horzcat(x, y), all_names, all_ext, ...
    'uniformoutput', false);
clear all_names all_ext


%% Read spreadsheet
ExcelPath = fullfile(root_dir,xl_fname);
[~,~,raw] = xlsread(ExcelPath,'Masking');
if isempty(raw)
    return;
end

%% Split into header and body
xlsx_head = raw(1,1:2);
fnames_old = raw(2:end, strcmpi(xlsx_head, OLD_LABEL));
fnames_new = raw(2:end, strcmpi(xlsx_head, NEW_LABEL));
clear raw

%% Waitbar
wb = waitbar(0, sprintf('Renaming %s%s...', ...
    fnames_old{1}, fnames_old{1}));
wb.Children.Title.Interpreter = 'none';
waitbar(0, wb, sprintf('Renaming %s...', fnames_old{1}));

%% Start renaming
n_files = numel(fnames_old);
notfound_count = 0;
for ii=1:n_files
    waitbar(ii/n_files, wb, sprintf('Renaming %s', fnames_old{ii}))
    
    % Search for file
    search_index = find(strcmp(all_fnames, fnames_old{ii}));
    current_path = all_paths(search_index);
    current_fname = all_fnames(search_index);
    % Handle unexpected results
    if isempty(search_index)
        notfound_count = notfound_count + 1;
        warning('%s not found', fnames_old{ii});
        continue;
    elseif numel(search_index) > 1
        warning('Duplicate %s in:', fnames_old{ii});
        for jj=1:numel(current_path)
            fprintf('%s\n', current_path{jj})
        end
        continue;
    end
    % Convert to char array
    current_path = current_path{1};
    current_fname = current_fname{1};
    
    
    % Rename
    if isequal(current_fname,'RenameCode.xlsm')
           continue;
    end
     
    
    [ok, msg] = movefile(...
        fullfile(current_path, fnames_old{ii}), ...
        fullfile(current_path, fnames_new{ii}));
    if ~ok
        warning(msg);
    end
    
    % Update progress
    waitbar(ii/n_files, wb)
end
close(wb)
% Summary of files, errors
fprintf('\nNumber of total files expected to rename: %g\n',n_files)
fprintf('Number of files not found: %g\n',notfound_count)
fprintf('Total number of files renamed: %g\n\n',n_files - notfound_count)
