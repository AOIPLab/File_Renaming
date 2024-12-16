%% Get file namess

clear all
clc

%% Imports
addpath('lib');

%% Constants
xl_fname = 'FileNames.xls';

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
writecell(all_fnames,ExcelPath,'Sheet',1,'Range','A1');

