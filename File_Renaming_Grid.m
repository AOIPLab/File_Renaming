%% Replace file names according to a provided spreadsheet of old and new 
% file names
% Not approved for AOIP use. Last validated: Never

%% Get spreadsheet
[fname_xlsx, path_xlsx] = uigetfile('*.xlsx', ...
    'Select spreadsheet with old and new file names');
if isnumeric(fname_xlsx)
    return;
end

%% Read spreadsheet
[~,~,raw] = xlsread(fullfile(path_xlsx, fname_xlsx));
if isempty(raw)
    return;
end

%% Split into header and body
xlsx_head = raw(1,:);
fnames_old = raw(2:end, 1);
fnames_new = raw(2:end, 2);

%% Waitbar
wb = waitbar(0, 'Renaming files...');

%% Start renaming
n_files = numel(fnames_old);
for ii=1:n_files
    [ok, msg] = movefile(...
        fullfile(path_xlsx, fnames_old{ii}), ...
        fullfile(path_xlsx, fnames_new{ii}));
    if ~ok
        warning(msg);
    end
    
    waitbar(ii/n_files, wb)
end
close(wb)

