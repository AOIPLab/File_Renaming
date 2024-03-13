%% Replace a substring in all files in a directory or all subdirectories
% Created - Alex Salmon - 2019.02.27
% Not yet approved for AOIP use. Last validated: yyyy.mm.dd

%% Imports
addpath('.\lib');

%% Get user input
prompt            = {'Desired directory path:',...
                     'Old string:',...
                     'New string:',...
                     'Search subfolders? (y/n)'};
dialog_name       =  'Replace a substring in all files';
defaults          = {'','','','n'};
opts.Resize       = 'on';

%% Loop until input is valid
input_valid = false;
while ~input_valid
    answers = inputdlg(prompt, dialog_name, 1, defaults, opts);
    if exist('mb', 'var') % Close message box
        close(mb);
    end
    if isempty(answers) % Cancel
        return;
    end
    
    % Reset error_found
    error_found = false;
    % Update defaults for quick fixes
    defaults = answers;

    % unpackaging input data
    warning_msg = '';
    root_dir            = answers{1};
    if exist(root_dir, 'dir') == 0
        warning_msg = 'Directory not found';
        error_found = true;
    end
    string_to_replace   = answers{2};
    desired_string      = answers{3};
    include_subdirs     = strtrim(answers{4});

    if ~error_found
        %% Search for files
        search = fullfile(root_dir, sprintf('*%s*', string_to_replace));
        if strcmpi(include_subdirs, 'y')
            f_list = subdir(search);
            % subdir outputs name as the ffname, fix that for the next step
            for ii=1:numel(f_list)
                [~,name,ext] = fileparts(f_list(ii).name);
                f_list(ii).name = [name,ext];
            end
        elseif strcmpi(include_subdirs, 'n')
            f_list = dir(search);
        else
            warning_msg = 'Input not recognized';
            error_found = true;
        end
        % Check if search comes up empty
        if isempty(f_list)
            warning_msg = 'No files found with specified substring';
            error_found = true;
        end
    end
    
    %% Validate input
    if ~error_found
        input_valid = true;
    else
        mb = msgbox(warning_msg, 'Error', 'Error');
    end
end

%% Waitbar
wb = waitbar(0, sprintf('Replacing substring "%s" as "%s"', ...
    string_to_replace, desired_string));
wb.Children.Title.Interpreter = 'none';

%% Rename files
for ii=1:numel(f_list)
    in_name     = f_list(ii).name;
    out_name    = strrep(in_name, string_to_replace, desired_string);
    in_ffname   = fullfile(f_list(ii).folder, in_name);
    out_ffname  = fullfile(f_list(ii).folder, out_name);
    % Rename w/ cmd line
    eval(sprintf('! ren "%s" "%s"', in_ffname, out_name));
    waitbar(ii/numel(f_list), wb)
end

%% Done
close(wb)









