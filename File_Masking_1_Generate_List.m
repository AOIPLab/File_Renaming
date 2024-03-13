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
SHEET_NAME = 'Masking';
xl_fname = 'RenameCode.xlsm';
% %% Get spreadsheet
% [xl_fname, xl_path] = uigetfile('*.xls*', ...
%     'Select spreadsheet with old and new file names', ...
%     'multiselect', 'off');
% if isnumeric(xl_fname)
%     return;
% end

%% Get root directory and do recursive file search
root_dir = uigetdir;
srch_results = subdir(root_dir);
srch_ffnames = {srch_results.name}';
[all_paths, all_names, all_ext] = cellfun(@fileparts, srch_ffnames, ...
    'uniformoutput', false);
all_fnames = cellfun(@(x, y) horzcat(x, y), all_names, all_ext, ...
    'uniformoutput', false);
clear all_names all_ext

%% Write to Spreadsheet
ExcelPath = fullfile(root_dir,xl_fname);
writecell(all_fnames,ExcelPath,'Sheet',1,'Range','A2');