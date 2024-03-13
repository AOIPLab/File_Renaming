%% Generate list of file names and extensions in an Excel sheet 
% Not approved for AOIP use. Last validated: Never
% Created - 2022.10.24 - Brian Higgins
% files and end of file summary report
%   Use recursive file search

%% Imports
addpath('lib');

%% Get root directory and do recursive file search
root_dir = uigetdir;
working_dir = regexp(root_dir,'\','split');
working_dir = char(working_dir(1,end));
srch_results = subdir(root_dir);
srch_ffnames = {srch_results.name}';
[all_paths, all_names, all_ext] = cellfun(@fileparts, srch_ffnames, ...
    'uniformoutput', false);
all_fnames = cellfun(@(x, y) horzcat(x, y), all_names, all_ext, ...
    'uniformoutput', false);
clear all_names all_ext

%% Write to Spreadsheet

xl_fname = [working_dir '.xlsm'];
ExcelPath = fullfile(root_dir,xl_fname);
writecell(all_fnames,ExcelPath,'Sheet',1,'Range','A2');