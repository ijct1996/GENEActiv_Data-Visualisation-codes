% ------------------------------------------------------------------------
% Script: Actograph Analysis and Summary for Parkinson’s Participants
% ------------------------------------------------------------------------
% PURPOSE:
%   Processes raw actigraphy `.xlsx` data to generate:
%     • Weekly participant activity profiles
%     • Weekly activity heatmaps
%     • Weekly daily activity and light-exposure bar charts
%     • Weekly low-activity call-outs
%     • A full-period light daily distribution area plot
%     • An Excel summary (“call-out sheet”) of key daily metrics + rhythm metrics,
%       including L5/M10 calculated on activity, plus a definitions sheet
%
% HOW TO RUN:
%   1. Run this script in MATLAB.
%   2. When prompted:
%        – Select the input Excel file containing actigraphy data.
%        – Select (or create) the output folder for saving results.
%   3. The script will process the data and save figures + an `.xlsx` summary.
%
% NOTE:
%   – Input Excel must have columns:
%       “Time stamp”       (format: yyyy-MM-dd HH:mm:ss:SSS)
%       “Sum of vector (SVMg)”  (activity)
%       “Light level (LUX)”      (light intensity)
%       “Button (1/0)”           (event markers)
% ------------------------------------------------------------------------

clc; clearvars; close all;

%% 1) Select and load the Excel file
[fileName, filePath] = uigetfile('*.xlsx','Select Excel file');
if isequal(fileName,0)
    error('File selection cancelled.');
end
fullFileName = fullfile(filePath,fileName);
dataTable    = readtable(fullFileName,'VariableNamingRule','preserve');

% Extract relevant columns
aActivity  = dataTable.("Sum of vector (SVMg)");   % raw activity
lLight     = dataTable.("Light level (LUX)");      % light intensity
bButton    = dataTable.("Button (1/0)");           % event markers
timestamps = datetime(dataTable.("Time stamp"),...
                 'InputFormat','yyyy-MM-dd HH:mm:ss:SSS');

%% 2) Bin into days × minutes
startTime         = dateshift(timestamps(1),'start','day');
minutesSinceStart = minutes(timestamps - startTime);
numDays           = ceil(days(timestamps(end) - startTime));
binEdges          = 0 : 1440 * numDays;
binIdx            = discretize(minutesSinceStart, binEdges);

% Build a datetime array for each “day 1, day 2, …”
dayDates = startTime + days(0:(numDays-1))';
weekDays = cellstr(datestr(dayDates,'dddd'));

% Preallocate binned data
binnedActivity = nan(numDays,1440);
binnedLight    = nan(numDays,1440);
binnedButton   = nan(numDays,1440);

% Populate bins
for d = 1:numDays
    sel    = binIdx > (d-1)*1440 & binIdx <= d*1440;
    relMin = binIdx(sel) - (d-1)*1440;
    binnedActivity(d,relMin) = aActivity(sel);
    binnedLight(d,relMin)    = lLight(sel);
    binnedButton(d,relMin)   = bButton(sel);
end

%% 3) Select output folder
outputFolder = uigetdir('','Select output folder for JPEGs');
if outputFolder == 0
    error('No folder selected.');
end
timeAxis = linspace(0,1440,1440);   % minutes for x-axis

%% 4) Weekly figures & low-activity call-outs
nWeeks = ceil(numDays/7);

for wk = 1:nWeeks
    daysIdx   = (wk-1)*7 + (1:7);
    daysIdx   = daysIdx(daysIdx <= numDays);
    nThisWeek = numel(daysIdx);
    
    % Prepare dd/mm labels (lowercase 'mm' = month)
    dateStrings = cellstr(datestr(dayDates(daysIdx),'dd/mm'));

    %% 4a) Participant Activity Profile (Day #)
    fig = figure('Color','w','Name',sprintf('Participant Activity Profile – Week %d',wk),...
                 'NumberTitle','off','Position',[100 100 1200 900],'Toolbar','none');
    for i = 1:nThisWeek
        d = daysIdx(i);
        ax = subplot(nThisWeek,1,i); hold(ax,'on');
        maskDark  = binnedLight(d,:) <= 1;
        maskLight = binnedLight(d,:) > 1;
        yTop = max(binnedActivity(d,:),[],'omitnan') * 1.2;
        if any(maskDark)
            xD = [timeAxis(maskDark), fliplr(timeAxis(maskDark))];
            yD = [zeros(1,sum(maskDark)), yTop*ones(1,sum(maskDark))];
            fill(ax, xD, yD, [0.9 0.9 0.9], 'EdgeColor','none','FaceAlpha',0.3);
        end
        if any(maskLight)
            xL = [timeAxis(maskLight), fliplr(timeAxis(maskLight))];
            yL = [zeros(1,sum(maskLight)), yTop*ones(1,sum(maskLight))];
            fill(ax, xL, yL, [0.9290 0.6940 0.1250], 'EdgeColor','none','FaceAlpha',0.3);
        end
        bar(ax, timeAxis, binnedActivity(d,:), 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
        ylim(ax, [0 yTop]); xlim(ax, [0 1440]);
        xticks(ax, 0:120:1440);
        xticklabels(ax, {'00:00','02:00','04:00','06:00','08:00','10:00',...
                        '12:00','14:00','16:00','18:00','20:00','22:00','00:00'});
        ylabel(ax, sprintf('Day %d', d), 'FontSize',12);
        set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
        if i == 1
            title(ax,'Participant Activity Profile','FontSize',16);
        end
        hold(ax,'off');
    end
    xlabel('Time of Day','FontSize',12);
    exportgraphics(fig, fullfile(outputFolder,sprintf('01_ActivityProfile_Week%d.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4a_dated) Participant Activity Profile (dd/mm)
    fig = figure('Color','w','Name',sprintf('Participant Activity Profile – Week %d (dated)',wk),...
                 'NumberTitle','off','Position',[100 100 1200 900],'Toolbar','none');
    for i = 1:nThisWeek
        d = daysIdx(i);
        ax = subplot(nThisWeek,1,i); hold(ax,'on');
        maskDark  = binnedLight(d,:) <= 1;
        maskLight = binnedLight(d,:) > 1;
        yTop = max(binnedActivity(d,:),[],'omitnan') * 1.2;
        if any(maskDark)
            xD = [timeAxis(maskDark), fliplr(timeAxis(maskDark))];
            yD = [zeros(1,sum(maskDark)), yTop*ones(1,sum(maskDark))];
            fill(ax, xD, yD, [0.9 0.9 0.9],'EdgeColor','none','FaceAlpha',0.3);
        end
        if any(maskLight)
            xL = [timeAxis(maskLight), fliplr(timeAxis(maskLight))];
            yL = [zeros(1,sum(maskLight)), yTop*ones(1,sum(maskLight))];
            fill(ax, xL, yL, [0.9290 0.6940 0.1250],'EdgeColor','none','FaceAlpha',0.3);
        end
        bar(ax, timeAxis, binnedActivity(d,:), 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
        ylim(ax, [0 yTop]); xlim(ax, [0 1440]);
        xticks(ax, 0:120:1440);
        xticklabels(ax, {'00:00','02:00','04:00','06:00','08:00','10:00',...
                        '12:00','14:00','16:00','18:00','20:00','22:00','00:00'});
        ylabel(ax, dateStrings{i}, 'FontSize',12);
        set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
        if i == 1
            title(ax,'Participant Activity Profile (dd/mm)','FontSize',16);
        end
        hold(ax,'off');
    end
    xlabel('Time of Day','FontSize',12);
    exportgraphics(fig, fullfile(outputFolder,sprintf('01_ActivityProfile_Week%d_dated.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4b) Activity Heatmap (Day #)
    fig = figure('Color','w','Name',sprintf('Activity Heatmap – Week %d',wk),...
                 'NumberTitle','off','Position',[100 100 1200 600],'Toolbar','none');
    blockData = binnedActivity(daysIdx,:);
    imagesc(timeAxis,1:nThisWeek,blockData);
    axis xy; set(gca,'YDir','reverse');
    caxis([prctile(blockData(~isnan(blockData)),5), prctile(blockData(~isnan(blockData)),95)]);
    colormap(parula);
    cb = colorbar; cb.Label.String = 'Activity (SVMg)';
    xlabel('Time of Day','FontSize',12); ylabel('Day','FontSize',12);
    xticks(0:360:1440); xticklabels({'00:00','06:00','12:00','18:00','24:00'});
    set(gca,'TickDir','out','FontSize',11,'Box','off');
    title('Activity Heatmap','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('02_Activity_Heatmap_Week%d.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4b_dated) Activity Heatmap (dd/mm)
    fig = figure('Color','w','Name',sprintf('Activity Heatmap – Week %d (dated)',wk),...
                 'NumberTitle','off','Position',[100 100 1200 600],'Toolbar','none');
    imagesc(timeAxis,1:nThisWeek,blockData);
    axis xy; set(gca,'YDir','reverse');
    caxis([prctile(blockData(~isnan(blockData)),5), prctile(blockData(~isnan(blockData)),95)]);
    colormap(parula);
    cb = colorbar; cb.Label.String = 'Activity (SVMg)';
    xlabel('Time of Day','FontSize',12);
    yticks(1:nThisWeek); yticklabels(dateStrings);
    xticks(0:360:1440); xticklabels({'00:00','06:00','12:00','18:00','24:00'});
    set(gca,'TickDir','out','FontSize',11,'Box','off');
    title('Activity Heatmap (dd/mm)','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('02_Activity_Heatmap_Week%d_dated.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4c) Daily Activity Bar Chart (Day #)
    fig = figure('Color','w','Name',sprintf('Daily Activity – Week %d',wk),'NumberTitle','off');
    da = nansum(binnedActivity(daysIdx,:),2);
    bar(1:nThisWeek,da,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
    xticks(1:nThisWeek);
    xticklabels(arrayfun(@(d)sprintf('Day %d',d),daysIdx,'UniformOutput',false));
    ylabel('Total Activity','FontSize',14); xlabel('Day','FontSize',14);
    set(gca,'TickDir','out','FontSize',12,'Box','off');
    title('Total Activity by Day','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('03_DailyActivity_Week%d.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4c_dated) Daily Activity Bar Chart (dd/mm)
    fig = figure('Color','w','Name',sprintf('Daily Activity – Week %d (dated)',wk),'NumberTitle','off');
    bar(1:nThisWeek,da,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
    xticks(1:nThisWeek); xticklabels(dateStrings);
    ylabel('Total Activity','FontSize',14); xlabel('Date','FontSize',14);
    set(gca,'TickDir','out','FontSize',12,'Box','off');
    title('Total Activity by Date (dd/mm)','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('03_DailyActivity_Week%d_dated.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4d) Daily Light Exposure Bar Chart (Day #)
    fig = figure('Color','w','Name',sprintf('Daily Light – Week %d',wk),'NumberTitle','off');
    dl = sum(binnedLight(daysIdx,:)>1,2)/60;
    bar(1:nThisWeek,dl,'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');
    xticks(1:nThisWeek);
    xticklabels(arrayfun(@(d)sprintf('Day %d',d),daysIdx,'UniformOutput',false));
    ylabel('Hours in Light','FontSize',14); xlabel('Day','FontSize',14);
    set(gca,'TickDir','out','FontSize',12,'Box','off');
    title('Daily Hours in Light','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('04_DailyLight_Week%d.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4d_dated) Daily Light Exposure Bar Chart (dd/mm)
    fig = figure('Color','w','Name',sprintf('Daily Light – Week %d (dated)',wk),'NumberTitle','off');
    bar(1:nThisWeek,dl,'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');
    xticks(1:nThisWeek); xticklabels(dateStrings);
    ylabel('Hours in Light','FontSize',14); xlabel('Date','FontSize',14);
    set(gca,'TickDir','out','FontSize',12,'Box','off');
    title('Daily Hours in Light by Date (dd/mm)','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('04_DailyLight_Week%d_dated.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4e) Low-Activity Call-outs (Day #)
    fig = figure('Color','w','Name',sprintf('Low Activity – Week %d',wk),'NumberTitle','off');
    bar(1:nThisWeek,da,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none'); hold on;
    lowIdx = find(da < (mean(da)-std(da)));
    for idx = lowIdx'
        text(idx, da(idx)+0.05*max(da), 'Low','Color','r','FontSize',12,'HorizontalAlign','center');
    end
    xticks(1:nThisWeek);
    xticklabels(arrayfun(@(d)sprintf('Day %d',d),daysIdx,'UniformOutput',false));
    ylabel('Total Activity','FontSize',14); xlabel('Day','FontSize',14);
    set(gca,'TickDir','out','FontSize',12,'Box','off');
    title('Low-Activity Call-outs','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('05_LowActivity_Week%d.jpg',wk)),'Resolution',600);
    close(fig);

    %% 4e_dated) Low-Activity Call-outs (dd/mm)
    fig = figure('Color','w','Name',sprintf('Low Activity – Week %d (dated)',wk),'NumberTitle','off');
    bar(1:nThisWeek,da,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none'); hold on;
    for idx = lowIdx'
        text(idx, da(idx)+0.05*max(da), 'Low','Color','r','FontSize',12,'HorizontalAlign','center');
    end
    xticks(1:nThisWeek); xticklabels(dateStrings);
    ylabel('Total Activity','FontSize',14); xlabel('Date','FontSize',14);
    set(gca,'TickDir','out','FontSize',12,'Box','off');
    title('Low-Activity Call-outs by Date (dd/mm)','FontSize',16);
    exportgraphics(fig, fullfile(outputFolder,sprintf('05_LowActivity_Week%d_dated.jpg',wk)),'Resolution',600);
    close(fig);
end

%% 5) Area Plot: Light Daily Distribution
hourOfDay      = hour(timestamps);
hourlyAvgLight = arrayfun(@(h) mean(lLight(hourOfDay==h),'omitnan'),0:23);

fig = figure('Color','w','Name','Light Daily Distribution','NumberTitle','off',...
             'Position',[100 100 900 500],'Toolbar','none');
area(0:23, hourlyAvgLight, 'FaceColor',[0.4 0.7608 0.6471]);
title('Light Daily Distribution','FontSize',16,'FontWeight','bold');
xlabel('Hour of Day (0 = Midnight)','FontSize',14);
ylabel('Average Light (LUX)','FontSize',14);
xlim([0 23]); xticks(0:1:23);
set(gca,'TickDir','out','FontSize',11,'Box','off');
exportgraphics(fig, fullfile(outputFolder,'06_LightDailyDistribution.jpg'),'Resolution',600);
close(fig);

%% 6) Call-Out Sheet for Quantitative Summary + Rhythm Metrics

% Core daily metrics
days       = (1:numDays)';
totalAct   = nansum(binnedActivity,2);
hoursLight = sum(binnedLight>1,2) / 60;

% Peak activity time
peakIdx  = arrayfun(@(d) find(binnedActivity(d,:)==max(binnedActivity(d,:)),1), days);
peakTime = timeAxis(peakIdx)';
peakDur  = minutes(peakTime); peakDur.Format = 'hh:mm:ss';
peakStr  = string(peakDur);

% Light start/end times
lightStartMin = nan(numDays,1);
lightEndMin   = nan(numDays,1);
for d = 1:numDays
    mask = binnedLight(d,:)>1;
    if any(mask)
        lightStartMin(d) = timeAxis(find(mask,1,'first'));
        lightEndMin(d)   = timeAxis(find(mask,1','last'));
    end
end
startDur = minutes(lightStartMin); startDur.Format = 'hh:mm:ss';
endDur   = minutes(lightEndMin);   endDur.Format   = 'hh:mm:ss';
startStr = string(startDur);
endStr   = string(endDur);

% L5/M10 on activity
L5_window  = 5*60;
M10_window = 10*60;
L5_start   = NaT(numDays,1);
M10_start  = NaT(numDays,1);
L5_mean    = NaN(numDays,1);
M10_mean   = NaN(numDays,1);

for d = 1:numDays
    signal = binnedActivity(d,:);
    conv5   = conv(signal, ones(1,L5_window)/L5_window,  'valid');
    conv10  = conv(signal, ones(1,M10_window)/M10_window,'valid');
    [L5_mean(d), idx5]   = min(conv5);
    [M10_mean(d), idx10] = max(conv10);
    L5_start(d)  = dayDates(d) + minutes(idx5-1);
    M10_start(d) = dayDates(d) + minutes(idx10-1);
end

L5_str  = string(timeofday(L5_start));   % hh:mm:ss
M10_str = string(timeofday(M10_start));  % hh:mm:ss

% Interdaily Stability (IS) & Intradaily Variability (IV)
M   = binnedActivity;
mh  = nanmean(M,1);
m   = nanmean(M(:));
IS  = (numDays * nansum((mh - m).^2)) / nansum((M(:) - m).^2);
x   = M(:); valid = ~isnan(x); nTot = sum(valid);
d2  = diff(x(valid));
IV  = (nTot * nansum(d2.^2)/(nTot-1)) / nansum((x(valid) - m).^2);

% Turn IS & IV into full-length columns
IS_col = [IS;  nan(numDays-1,1)];
IV_col = [IV;  nan(numDays-1,1)];

% Normative ranges
IS_norm     = "0.58-0.73";
IV_norm     = "0.56-0.77";
IS_norm_col = repmat(IS_norm, numDays,1);
IV_norm_col = repmat(IV_norm, numDays,1);

% Write to Excel with two sheets
excelFile = fullfile(outputFolder,'07_Participant_results.xlsx');

% --- Sheet 1: Summary ---
Callout = table( ...
    dayDates, weekDays, days, totalAct, hoursLight, peakStr, startStr, endStr, ...
    L5_str, L5_mean, M10_str, M10_mean, ...
    IS_col, IV_col, ...
    IS_norm_col, IV_norm_col, ...
    'VariableNames',{ ...
      'Date','Weekday','Day','TotalActivity','HoursInLight', ...
      'PeakActivityTime','LightStartTime','LightEndTime', ...
      'L5_StartTime','L5_Mean','M10_StartTime','M10_Mean', ...
      'InterdailyStability','IntradailyVariability', ...
      'IS_NormalRange','IV_NormalRange' ...
    } ...
);
writetable(Callout, excelFile, 'Sheet', 'Summary');

% --- Sheet 2: Definitions ---
headers   = {'Term','Definition','Interpretation'};
termsData = {
  'Date',                 'Calendar date of recording',                          'Aligns metrics to calendar days';
  'Weekday',              'Day of the week',                                     'Distinguishes weekday vs weekend';
  'Day',                  'Sequential day index',                                'Day number since start';
  'TotalActivity',        'Sum of minute-by-minute activity',                    'Overall daily movement';
  'HoursInLight',         'Total hours with LUX >1',                             'Duration of light exposure';
  'PeakActivityTime',     'Time of maximum minute activity',                     'Indicates peak movement time';
  'LightStartTime',       'Time of first LUX >1',                                'Onset of light exposure';
  'LightEndTime',         'Time of last LUX >1',                                 'End of light exposure';
  'L5_StartTime',         'The clock time at which the consecutive 5-hour window of lowest mean activity begins',              'Marks the onset of the daily rest-activity trough. A later or more variable L5 start can signal disrupted rest';
  'L5_Mean',              'The average minute-by-minute activity count across that lowest 5-hour window',                      'Quantifies the depth of the rest-activity trough. Lower values imply deeper or more consolidated rest periods';
  'M10_StartTime',        'The clock time at which the consecutive 10-hour window of highest mean activity begins',            'Identifies when the most sustained active period starts. Shifts earlier or later can indicate advanced/delayed phase';
  'M10_Mean',             'The average minute-by-minute activity count across that highest 10-hour window',                     'Reflects overall vigour during peak activity. Higher values suggest stronger or more sustained daytime activity';
  'InterdailyStability',  'Strength of 24-h rhythm',                             'Higher = more stable';
  'IntradailyVariability','Fragmentation of activity rhythm',                    'Higher = more fragmented';
  'IS_NormalRange',       'Normative IS range (0.58–0.73)',                      'Typical healthy range';
  'IV_NormalRange',       'Normative IV range (0.56–0.77)',                      'Typical healthy range'
};
writecell([headers; termsData], excelFile, 'Sheet', 'Definitions');

%% 7) Bundle all figures into a PowerPoint presentation
import mlreportgen.ppt.*;

pptFile = fullfile(outputFolder,'AllFigures_Report.pptx');
ppt     = Presentation(pptFile);
open(ppt);

jpgs = dir(fullfile(outputFolder,'*.jpg'));
for i = 1:numel(jpgs)
    slide = add(ppt,'Title and Content');
    replace(slide,'Title', erase(jpgs(i).name,'.jpg'));
    pic = Picture(fullfile(outputFolder,jpgs(i).name));
    replace(slide,'Content',pic);
end

close(ppt);
close all;