% ------------------------------------------------------------------------
% Script: Actograph Analysis and Summary for Parkinson’s Participants
% ------------------------------------------------------------------------
% PURPOSE:
%   Processes raw actigraphy `.xlsx` data to generate:
%     • Weekly participant activity profiles
%     • Weekly activity heatmaps
%     • Weekly daily activity and light‐exposure bar charts
%     • Weekly low‐activity call‐outs
%     • A full‐period light daily distribution area plot
%     • An Excel workbook with two sheets:
%         – Summary of key daily metrics and rhythm metrics, including L5/M10 on activity
%         – Definitions of all terms with interpretation guidance
%
% HOW TO RUN:
%   1. Run this script in MATLAB.
%   2. When prompted:
%        – Select the input Excel file containing the actigraphy data.
%        – Choose whether to plot only complete 24‐hour days or all days.
%        – Select or create the output folder for saving JPEG images.
%   3. The script will process the data and save figures and an Excel workbook.
%
% NOTE:
%   Input Excel must have columns:
%     “Time stamp”       (format: yyyy‐MM‐dd HH:mm:ss:SSS)
%     “Sum of vector (SVMg)”  (activity)
%     “Light level (LUX)”      (light intensity)
%     “Button (1/0)”           (event markers)
% ------------------------------------------------------------------------

clc; clearvars; close all;

%% 1) Select and load the Excel file
[fileName, filePath] = uigetfile('*.xlsx', 'Select the input Excel file');
if isequal(fileName, 0)
    error('File selection cancelled.');
end
fullFileName = fullfile(filePath, fileName);
dataTable    = readtable(fullFileName, 'VariableNamingRule', 'preserve');

% Extract relevant columns
activityCounts  = dataTable.("Sum of vector (SVMg)");   % raw activity data
lightLevels     = dataTable.("Light level (LUX)");      % light intensity
buttonEvents    = dataTable.("Button (1/0)");           % event markers
timestamps      = datetime(dataTable.("Time stamp"), ...
                  'InputFormat', 'yyyy-MM-dd HH:mm:ss:SSS');

%% 2) Bin data into days by minutes
firstDayStart       = dateshift(timestamps(1), 'start', 'day');
elapsedMinutes      = minutes(timestamps - firstDayStart);
totalDays           = ceil(days(timestamps(end) - firstDayStart));
binEdges            = 0 : 1440 * totalDays;
minuteBinIndices    = discretize(elapsedMinutes, binEdges);

% Build arrays of dates
dayStartDates = firstDayStart + days(0:(totalDays-1))';
weekDayNames  = cellstr(datestr(dayStartDates, 'dddd'));

% Preallocate binned matrices
binnedActivity = nan(totalDays, 1440);
binnedLight    = nan(totalDays, 1440);
binnedButton   = nan(totalDays, 1440);

% Populate binned matrices
for dayIndex = 1:totalDays
    selector     = minuteBinIndices > (dayIndex-1)*1440 & minuteBinIndices <= dayIndex*1440;
    relativeMins  = minuteBinIndices(selector) - (dayIndex-1)*1440;
    binnedActivity(dayIndex, relativeMins) = activityCounts(selector);
    binnedLight(dayIndex, relativeMins)    = lightLevels(selector);
    binnedButton(dayIndex, relativeMins)   = buttonEvents(selector);
end

%% 3) Prompt user for data range choice
choice = questdlg( ...
    'Plot only complete 24‐hour days, or all available days?', ...
    'Select Data Range', ...
    'Complete days','All days','All days' ...
    );
if strcmp(choice, 'Complete days')
    isCompleteDay = all(~isnan(binnedActivity), 2);
    dayStartDates = dayStartDates(isCompleteDay);
    weekDayNames  = weekDayNames(isCompleteDay);
    binnedActivity = binnedActivity(isCompleteDay, :);
    binnedLight    = binnedLight(isCompleteDay, :);
    binnedButton   = binnedButton(isCompleteDay, :);
    totalDays      = sum(isCompleteDay);
else
    dayStartDates = dayStartDates;
    weekDayNames  = weekDayNames;
    % totalDays remains unchanged
end

%% 4) Select output folder for JPEG images
outputFolder = uigetdir('', 'Select output folder for JPEG images');
if outputFolder == 0
    error('No folder selected for output.');
end
timeAxisMinutes = linspace(0, 1440, 1440);  % horizontal axis in minutes

%% 5) Generate weekly figures and call‐outs
numberOfWeeks = ceil(totalDays / 7);

for weekIndex = 1:numberOfWeeks
    dayIndices     = (weekIndex-1)*7 + (1:7);
    dayIndices     = dayIndices(dayIndices <= totalDays);
    daysThisWeek   = numel(dayIndices);
    
    % Prepare dd/mm labels for this week
    dateLabels = cellstr(datestr(dayStartDates(dayIndices), 'dd/mm'));

    %% 5a) Participant Activity Profile (Day number)
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Participant Activity Profile – Week %d', weekIndex), ...
                 'NumberTitle', 'off', 'Position', [100 100 1200 900], ...
                 'Toolbar', 'none');
    for i = 1:daysThisWeek
        dayIdx = dayIndices(i);
        ax = subplot(daysThisWeek, 1, i); hold(ax, 'on');
        darkMask  = binnedLight(dayIdx, :) <= 1;
        lightMask = binnedLight(dayIdx, :) > 1;
        yMax      = max(binnedActivity(dayIdx, :), [], 'omitnan') * 1.2;
        if any(darkMask)
            xDark = [timeAxisMinutes(darkMask), fliplr(timeAxisMinutes(darkMask))];
            yDark = [zeros(1, sum(darkMask)), yMax * ones(1, sum(darkMask))];
            fill(ax, xDark, yDark, [0.9 0.9 0.9], 'EdgeColor', 'none', 'FaceAlpha', 0.3);
        end
        if any(lightMask)
            xLight = [timeAxisMinutes(lightMask), fliplr(timeAxisMinutes(lightMask))];
            yLight = [zeros(1, sum(lightMask)), yMax * ones(1, sum(lightMask))];
            fill(ax, xLight, yLight, [0.9290 0.6940 0.1250], 'EdgeColor', 'none', 'FaceAlpha', 0.3);
        end
        bar(ax, timeAxisMinutes, binnedActivity(dayIdx, :), ...
            'FaceColor', [0 0.4470 0.7410], 'EdgeColor', 'none');
        ylim(ax, [0 yMax]); xlim(ax, [0 1440]);
        xticks(ax, 0:120:1440);
        xticklabels(ax, { ...
            '00:00','02:00','04:00','06:00','08:00','10:00', ...
            '12:00','14:00','16:00','18:00','20:00','22:00','00:00' });
        ylabel(ax, sprintf('Day %d', dayIdx), 'FontSize', 12);
        set(ax, 'TickDir', 'out', 'YTick', [], 'Box', 'off', 'FontSize', 11);
        if i == 1
            title(ax, 'Participant Activity Profile', 'FontSize', 16);
        end
        hold(ax, 'off');
    end
    xlabel('Time of Day', 'FontSize', 12);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('01_ActivityProfile_Week%d.jpg', weekIndex)), 'Resolution', 600);
    close(fig);

    %% 5b) Participant Activity Profile (dated)
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Participant Activity Profile – Week %d (dated)', weekIndex), ...
                 'NumberTitle', 'off', 'Position', [100 100 1200 900], ...
                 'Toolbar', 'none');
    for i = 1:daysThisWeek
        dayIdx = dayIndices(i);
        ax = subplot(daysThisWeek, 1, i); hold(ax, 'on');
        darkMask  = binnedLight(dayIdx, :) <= 1;
        lightMask = binnedLight(dayIdx, :) > 1;
        yMax      = max(binnedActivity(dayIdx, :), [], 'omitnan') * 1.2;
        if any(darkMask)
            xDark = [timeAxisMinutes(darkMask), fliplr(timeAxisMinutes(darkMask))];
            yDark = [zeros(1, sum(darkMask)), yMax * ones(1, sum(darkMask))];
            fill(ax, xDark, yDark, [0.9 0.9 0.9], 'EdgeColor', 'none', 'FaceAlpha', 0.3);
        end
        if any(lightMask)
            xLight = [timeAxisMinutes(lightMask), fliplr(timeAxisMinutes(lightMask))];
            yLight = [zeros(1, sum(lightMask)), yMax * ones(1, sum(lightMask))];
            fill(ax, xLight, yLight, [0.9290 0.6940 0.1250], 'EdgeColor', 'none', 'FaceAlpha', 0.3);
        end
        bar(ax, timeAxisMinutes, binnedActivity(dayIdx, :), ...
            'FaceColor', [0 0.4470 0.7410], 'EdgeColor', 'none');
        ylim(ax, [0 yMax]); xlim(ax, [0 1440]);
        xticks(ax, 0:120:1440);
        xticklabels(ax, { ...
            '00:00','02:00','04:00','06:00','08:00','10:00', ...
            '12:00','14:00','16:00','18:00','20:00','22:00','00:00' });
        ylabel(ax, dateLabels{i}, 'FontSize', 12);
        set(ax, 'TickDir', 'out', 'YTick', [], 'Box', 'off', 'FontSize', 11);
        if i == 1
            title(ax, 'Participant Activity Profile (dd/mm)', 'FontSize', 16);
        end
        hold(ax, 'off');
    end
    xlabel('Time of Day', 'FontSize', 12);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('01_ActivityProfile_Week%d_dated.jpg', weekIndex)), 'Resolution', 600);
    close(fig);

    %% 5c) Activity Heatmap and dated version
    blockData = binnedActivity(dayIndices, :);
    % heatmap with day numbers
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Activity Heatmap – Week %d', weekIndex), ...
                 'NumberTitle', 'off', 'Position', [100 100 1200 600], ...
                 'Toolbar', 'none');
    imagesc(timeAxisMinutes, 1:daysThisWeek, blockData);
    axis xy; set(gca, 'YDir', 'reverse');
    caxis([prctile(blockData(~isnan(blockData)),5), ...
           prctile(blockData(~isnan(blockData)),95)]);
    colormap(parula);
    cb = colorbar; cb.Label.String = 'Activity (SVMg)';
    xlabel('Time of Day', 'FontSize', 12); ylabel('Day', 'FontSize', 12);
    xticks(0:360:1440); xticklabels({'00:00','06:00','12:00','18:00','24:00'});
    set(gca, 'TickDir', 'out', 'FontSize', 11, 'Box', 'off');
    title('Activity Heatmap', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('02_Activity_Heatmap_Week%d.jpg', weekIndex)), 'Resolution', 600);
    close(fig);
    % dated heatmap
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Activity Heatmap – Week %d (dated)', weekIndex), ...
                 'NumberTitle', 'off', 'Position', [100 100 1200 600], ...
                 'Toolbar', 'none');
    imagesc(timeAxisMinutes, 1:daysThisWeek, blockData);
    axis xy; set(gca, 'YDir', 'reverse');
    caxis([prctile(blockData(~isnan(blockData)),5), ...
           prctile(blockData(~isnan(blockData)),95)]);
    colormap(parula);
    cb = colorbar; cb.Label.String = 'Activity (SVMg)';
    xlabel('Time of Day', 'FontSize', 12);
    yticks(1:daysThisWeek); yticklabels(dateLabels);
    xticks(0:360:1440); xticklabels({'00:00','06:00','12:00','18:00','24:00'});
    set(gca, 'TickDir', 'out', 'FontSize', 11, 'Box', 'off');
    title('Activity Heatmap (dd/mm)', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('02_Activity_Heatmap_Week%d_dated.jpg', weekIndex)), 'Resolution', 600);
    close(fig);

    %% 5d) Daily Activity Bar Chart and dated version
    dailyTotals = nansum(binnedActivity(dayIndices, :), 2);
    % by day number
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Daily Activity – Week %d', weekIndex), ...
                 'NumberTitle', 'off');
    bar(1:daysThisWeek, dailyTotals, 'FaceColor', [0 0.4470 0.7410], 'EdgeColor', 'none');
    xticks(1:daysThisWeek);
    xticklabels(arrayfun(@(d) sprintf('Day %d', d), dayIndices, 'UniformOutput', false));
    ylabel('Total Activity', 'FontSize', 14); xlabel('Day', 'FontSize', 14);
    set(gca, 'TickDir', 'out', 'FontSize', 12, 'Box', 'off');
    title('Total Activity by Day', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('03_DailyActivity_Week%d.jpg', weekIndex)), 'Resolution', 600);
    close(fig);
    % by date
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Daily Activity – Week %d (dated)', weekIndex), ...
                 'NumberTitle', 'off');
    bar(1:daysThisWeek, dailyTotals, 'FaceColor', [0 0.4470 0.7410], 'EdgeColor', 'none');
    xticks(1:daysThisWeek); xticklabels(dateLabels);
    ylabel('Total Activity', 'FontSize', 14); xlabel('Date', 'FontSize', 14);
    set(gca, 'TickDir', 'out', 'FontSize', 12, 'Box', 'off');
    title('Total Activity by Date (dd/mm)', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('03_DailyActivity_Week%d_dated.jpg', weekIndex)), 'Resolution', 600);
    close(fig);

    %% 5e) Daily Light‐Exposure Bar Chart and dated version
    hoursInLight = sum(binnedLight(dayIndices, :) > 1, 2) / 60;
    % by day number
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Daily Light – Week %d', weekIndex), ...
                 'NumberTitle', 'off');
    bar(1:daysThisWeek, hoursInLight, 'FaceColor', [0.8500 0.3250 0.0980], 'EdgeColor', 'none');
    xticks(1:daysThisWeek);
    xticklabels(arrayfun(@(d) sprintf('Day %d', d), dayIndices, 'UniformOutput', false));
    ylabel('Hours in Light', 'FontSize', 14); xlabel('Day', 'FontSize', 14);
    set(gca, 'TickDir', 'out', 'FontSize', 12, 'Box', 'off');
    title('Daily Hours in Light', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('04_DailyLight_Week%d.jpg', weekIndex)), 'Resolution', 600);
    close(fig);
    % by date
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Daily Light – Week %d (dated)', weekIndex), ...
                 'NumberTitle', 'off');
    bar(1:daysThisWeek, hoursInLight, 'FaceColor', [0.8500 0.3250 0.0980], 'EdgeColor', 'none');
    xticks(1:daysThisWeek); xticklabels(dateLabels);
    ylabel('Hours in Light', 'FontSize', 14); xlabel('Date', 'FontSize', 14);
    set(gca, 'TickDir', 'out', 'FontSize', 12, 'Box', 'off');
    title('Daily Hours in Light by Date (dd/mm)', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('04_DailyLight_Week%d_dated.jpg', weekIndex)), 'Resolution', 600);
    close(fig);

    %% 5f) Low‐Activity Call‐outs and dated version
    % by day number
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Low Activity – Week %d', weekIndex), ...
                 'NumberTitle', 'off');
    bar(1:daysThisWeek, dailyTotals, 'FaceColor', [0 0.4470 0.7410], 'EdgeColor', 'none');
    hold on;
    lowThreshold = mean(dailyTotals) - std(dailyTotals);
    lowIndices   = find(dailyTotals < lowThreshold);
    for idx = lowIndices'
        text(idx, dailyTotals(idx) + 0.05*max(dailyTotals), 'Low', ...
             'Color', 'r', 'FontSize', 12, 'HorizontalAlign', 'center');
    end
    xticks(1:daysThisWeek);
    xticklabels(arrayfun(@(d) sprintf('Day %d', d), dayIndices, 'UniformOutput', false));
    ylabel('Total Activity', 'FontSize', 14); xlabel('Day', 'FontSize', 14);
    set(gca, 'TickDir', 'out', 'FontSize', 12, 'Box', 'off');
    title('Low-Activity Call-outs', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('05_LowActivity_Week%d.jpg', weekIndex)), 'Resolution', 600);
    close(fig);
    % by date
    fig = figure('Color', 'w', ...
                 'Name', sprintf('Low Activity – Week %d (dated)', weekIndex), ...
                 'NumberTitle', 'off');
    bar(1:daysThisWeek, dailyTotals, 'FaceColor', [0 0.4470 0.7410], 'EdgeColor', 'none');
    hold on;
    for idx = lowIndices'
        text(idx, dailyTotals(idx) + 0.05*max(dailyTotals), 'Low', ...
             'Color', 'r', 'FontSize', 12, 'HorizontalAlign', 'center');
    end
    xticks(1:daysThisWeek); xticklabels(dateLabels);
    ylabel('Total Activity', 'FontSize', 14); xlabel('Date', 'FontSize', 14);
    set(gca, 'TickDir', 'out', 'FontSize', 12, 'Box', 'off');
    title('Low-Activity Call-outs by Date (dd/mm)', 'FontSize', 16);
    exportgraphics(fig, fullfile(outputFolder, ...
        sprintf('05_LowActivity_Week%d_dated.jpg', weekIndex)), 'Resolution', 600);
    close(fig);
end

%% 6) Full‐period light daily distribution area plot
hourOfDay      = hour(timestamps);
hourlyAverage  = arrayfun(@(h) mean(lightLevels(hourOfDay==h), 'omitnan'), 0:23);

fig = figure('Color', 'w', 'Name', 'Light Daily Distribution', ...
             'NumberTitle', 'off', 'Position', [100 100 900 500], 'Toolbar', 'none');
area(0:23, hourlyAverage, 'FaceColor', [0.4 0.7608 0.6471]);
title('Light Daily Distribution', 'FontSize', 16, 'FontWeight', 'bold');
xlabel('Hour of Day (0 = Midnight)', 'FontSize', 14);
ylabel('Average Light (LUX)', 'FontSize', 14);
xlim([0 23]); xticks(0:1:23);
set(gca, 'TickDir', 'out', 'FontSize', 11, 'Box', 'off');
exportgraphics(fig, fullfile(outputFolder, '06_LightDailyDistribution.jpg'), 'Resolution', 600);
close(fig);

%% 7) Call‐Out Sheet for Quantitative Summary + Rhythm Metrics

% Daily metrics
dayIndicesAll = (1:totalDays)';
totalActivity = nansum(binnedActivity, 2);
hoursInLight  = sum(binnedLight>1, 2) / 60;

% Peak activity time
peakIndices   = arrayfun(@(d) find(binnedActivity(d,:)==max(binnedActivity(d,:)), 1), dayIndicesAll);
peakTimes     = timeAxisMinutes(peakIndices)';
peakDurations = minutes(peakTimes); peakDurations.Format = 'hh:mm:ss';
peakStrings   = string(peakDurations);

% Light start and end times
lightStartMin = nan(totalDays,1);
lightEndMin   = nan(totalDays,1);
for d = 1:totalDays
    mask = binnedLight(d,:) > 1;
    if any(mask)
        lightStartMin(d) = timeAxisMinutes(find(mask,1,'first'));
        lightEndMin(d)   = timeAxisMinutes(find(mask,1,'last'));
    end
end
startDurations = minutes(lightStartMin); startDurations.Format = 'hh:mm:ss';
endDurations   = minutes(lightEndMin);   endDurations.Format   = 'hh:mm:ss';
startStrings   = string(startDurations);
endStrings     = string(endDurations);

% L5 and M10 on activity
L5WindowMinutes  = 5 * 60;
M10WindowMinutes = 10 * 60;
L5StartTimes     = NaT(totalDays,1);
M10StartTimes    = NaT(totalDays,1);
L5Means          = nan(totalDays,1);
M10Means         = nan(totalDays,1);

for d = 1:totalDays
    signal = binnedActivity(d,:);
    rolling5  = conv(signal, ones(1,L5WindowMinutes)/L5WindowMinutes,  'valid');
    rolling10 = conv(signal, ones(1,M10WindowMinutes)/M10WindowMinutes,'valid');
    [L5Means(d), idxL5]   = min(rolling5);
    [M10Means(d), idxM10] = max(rolling10);
    L5StartTimes(d)  = dayStartDates(d) + minutes(idxL5-1);
    M10StartTimes(d) = dayStartDates(d) + minutes(idxM10-1);
end

L5Strings  = string(timeofday(L5StartTimes));   % hh:mm:ss
M10Strings = string(timeofday(M10StartTimes));  % hh:mm:ss

% Interdaily Stability and Intradaily Variability
MMatrix = binnedActivity;
meanHourly = nanmean(MMatrix, 1);
grandMean  = nanmean(MMatrix(:));
ISMetric   = (totalDays * nansum((meanHourly - grandMean).^2)) / nansum((MMatrix(:) - grandMean).^2);
flatData   = MMatrix(:);
validData  = ~isnan(flatData);
diffs      = diff(flatData(validData));
IVMetric   = (sum(validData) * nansum(diffs.^2)/(sum(validData)-1)) / nansum((flatData(validData)-grandMean).^2);

% Convert IS and IV to full-length columns
ISColumn = [ISMetric;    nan(totalDays-1,1)];
IVColumn = [IVMetric;    nan(totalDays-1,1)];

% Normative ranges columns
ISRange      = "0.58-0.73";
IVRange      = "0.56-0.77";
ISRangeColumn = repmat(ISRange, totalDays,1);
IVRangeColumn = repmat(IVRange, totalDays,1);

% Write to Excel with two sheets
excelWorkbook = fullfile(outputFolder, '07_Participant_results.xlsx');

% Sheet 1: Summary
SummaryTable = table( ...
    dayStartDates, weekDayNames, dayIndicesAll, totalActivity, hoursInLight, ...
    peakStrings, startStrings, endStrings, ...
    L5Strings, L5Means, M10Strings, M10Means, ...
    ISColumn, IVColumn, ...
    ISRangeColumn, IVRangeColumn, ...
    'VariableNames', { ...
      'Date', 'Weekday', 'Day', 'TotalActivity', 'HoursInLight', ...
      'PeakActivityTime', 'LightStartTime', 'LightEndTime', ...
      'L5_StartTime', 'L5_Mean', 'M10_StartTime', 'M10_Mean', ...
      'InterdailyStability', 'IntradailyVariability', ...
      'IS_NormalRange', 'IV_NormalRange' } ...
);
writetable(SummaryTable, excelWorkbook, 'Sheet', 'Summary');

% Sheet 2: Definitions
DefinitionHeaders = {'Term','Definition','Interpretation'};
DefinitionData = {
  'Date',                 'Calendar date of recording',                          'Aligns metrics to calendar days';
  'Weekday',              'Day of the week',                                     'Distinguishes weekday vs weekend';
  'Day',                  'Sequential day index',                                'Day number since start';
  'TotalActivity',        'Sum of minute‐by‐minute activity',                    'Overall daily movement';
  'HoursInLight',         'Total hours during which light level is greater than one LUX', 'Duration of light exposure';
  'PeakActivityTime',     'Clock time of maximum minute activity',               'Indicates peak movement time';
  'LightStartTime',       'Clock time of first minute exceeding one LUX',        'Onset of daily light exposure';
  'LightEndTime',         'Clock time of last minute exceeding one LUX',         'End of daily light exposure';
  'L5_StartTime',         'Clock time when the five‐hour window of lowest mean activity begins', 'Onset of rest‐activity trough; later or variable times may indicate disturbed rest';
  'L5_Mean',              'Average activity counts during that lowest five‐hour window', 'Depth of rest‐activity trough; lower values indicate more consolidated rest';
  'M10_StartTime',        'Clock time when the ten‐hour window of highest mean activity begins', 'Onset of most sustained active period; shifts may indicate phase changes';
  'M10_Mean',             'Average activity counts during that highest ten‐hour window', 'Strength of daytime activity; higher values indicate more vigorous activity';
  'InterdailyStability',  'Measure of consistency of activity pattern across days', 'Higher values indicate more stable daily rhythm';
  'IntradailyVariability','Measure of fragmentation in the activity rhythm within days', 'Higher values indicate more fragmented activity';
  'IS_NormalRange',       'Typical range for Interdaily Stability (0.58 to 0.73)',    'Healthy adult reference range';
  'IV_NormalRange',       'Typical range for Intradaily Variability (0.56 to 0.77)', 'Healthy adult reference range'
};
writecell([DefinitionHeaders; DefinitionData], excelWorkbook, 'Sheet', 'Definitions');

%% 8) Bundle all JPEG figures into a PowerPoint presentation
import mlreportgen.ppt.*;

presentationFile = fullfile(outputFolder, 'AllFigures_Report.pptx');
presentation = Presentation(presentationFile);
open(presentation);

jpegFiles = dir(fullfile(outputFolder, '*.jpg'));
for i = 1:numel(jpegFiles)
    slide = add(presentation, 'Title and Content');
    replace(slide, 'Title', erase(jpegFiles(i).name, '.jpg'));
    image = Picture(fullfile(outputFolder, jpegFiles(i).name));
    replace(slide, 'Content', image);
end

close(presentation);
close all;
