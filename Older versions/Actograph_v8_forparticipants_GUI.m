function Actograph_v8_forparticipants_GUI
% -------------------------------------------------------------------------
% Actograph Analysis and Summary GUI for GENEActiv Participants
% -------------------------------------------------------------------------
% A self-contained MATLAB GUI for processing wrist-worn actigraphy data
% exported from GENEActiv and pasted into an Excel workbook.
%
% Expected input workbook
%   - Preferably contains a sheet named "RawData"
%   - If "RawData" is not present, the first sheet is used
%   - Required columns (matched robustly, ignoring whitespace and case):
%       * Time stamp
%       * Sum of vector (SVMg)
%       * Light level (LUX)
%       * Temperature
%
% Outputs (selected in GUI)
%   - High-resolution JPGs (weekly summaries and optional daily light tracker)
%   - Excel workbook with Summary, Metrics, Definitions
%   - Optional PowerPoint compiling all JPGs including subfolders
%
% Robustness choices in this version
%   - Keeps binning logic close to the original version (no timetable/retime)
%   - Prefers "RawData" sheet but falls back safely
%   - Header matching ignores extra whitespace (handles e.g., "Temperature ")
%   - Timestamp parsing tries multiple formats and drops invalid rows
%   - L5/M10 computed using NaN-robust sliding means (requires enough valid bins)
%   - IS/IV computed on hourly means derived from the binned daily matrix
%   - PowerPoint export includes JPGs in subfolders (e.g., DailyLightTracker)
%   - Status updates and a progress dialog report what is happening throughout
%
% Requirements
%   - MATLAB R2019a or newer
%   - Report Generator toolbox only if PowerPoint export is selected
%
% © 2025-2026 [Isaiah J Ting, Lall Lab]
% -------------------------------------------------------------------------

    clc; clearvars; close all;

    % ---------------------------------------------------------------------
    % Universal figure size and export defaults
    % ---------------------------------------------------------------------
    set(0, ...
        'DefaultFigureUnits','pixels', ...
        'DefaultFigurePosition',[100 100 1200 600], ...
        'DefaultFigurePaperUnits','inches', ...
        'DefaultFigurePaperPositionMode','auto');

    % Create UI figure
    fig = uifigure('Name','Actograph Analysis','Position',[300 200 600 580]);

    % Input file selection
    uilabel(fig,'Position',[20 540 70 22],'Text','Input file:');
    txtInputFile = uieditfield(fig,'text','Position',[100 540 340 22]);
    uibutton(fig,'push','Position',[450 540 120 22],'Text','Browse...', ...
        'ButtonPushedFcn',@selectInputFile);

    % Global light threshold
    uilabel(fig,'Position',[20 500 140 22],'Text','Light threshold (lux):');
    numThreshold = uieditfield(fig,'numeric','Position',[160 500 80 22], ...
        'Value',5,'Limits',[0 Inf],'RoundFractionalValues',true);

    % Complete days only checkbox
    chkCompleteOnly = uicheckbox(fig,'Position',[20 460 280 22], ...
        'Text','Include only complete 24-hour days');

    % Axis style dropdown
    uilabel(fig,'Position',[20 420 80 22],'Text','Axis style:');
    ddlAxisStyle = uidropdown(fig,'Position',[100 420 200 22], ...
        'Items',{'Days only','Dated only','Both'},'Value','Both');

    % Outputs selection panel
    pnlOutputs = uipanel(fig,'Title','Select Outputs','Position',[20 260 560 140]);

    chkAll             = uicheckbox(pnlOutputs,'Position',[10 100 120 20], ...
                          'Text','All outputs','Value',true,'ValueChangedFcn',@toggleAllOutputs);

    chkActivityProf    = uicheckbox(pnlOutputs,'Position',[10 75 200 20], ...
                          'Text','Activity profiles','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkDailyActivity   = uicheckbox(pnlOutputs,'Position',[10 50 220 20], ...
                          'Text','Daily activity bar charts','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkLowActivity     = uicheckbox(pnlOutputs,'Position',[10 25 220 20], ...
                          'Text','Low-activity call-outs','Value',true,'ValueChangedFcn',@syncAllCheckbox);

    chkHeatmap         = uicheckbox(pnlOutputs,'Position',[220 75 200 20], ...
                          'Text','Activity heatmaps','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkTempProfiles    = uicheckbox(pnlOutputs,'Position',[220 50 200 20], ...
                          'Text','Temperature profiles','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkCombined        = uicheckbox(pnlOutputs,'Position',[220 25 320 20], ...
                          'Text','Combined profile (Activity, Light, Temp)','Value',false,'ValueChangedFcn',@syncAllCheckbox);

    chkLightTracker    = uicheckbox(pnlOutputs,'Position',[420 75 200 20], ...
                          'Text','Daily light tracker','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkLightDist       = uicheckbox(pnlOutputs,'Position',[420 50 200 20], ...
                          'Text','Light distribution','Value',true,'ValueChangedFcn',@syncAllCheckbox);

    chkExcelSummary    = uicheckbox(pnlOutputs,'Position',[10 0 200 20], ...
                          'Text','Excel summary','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkExcelMetrics    = uicheckbox(pnlOutputs,'Position',[220 0 200 20], ...
                          'Text','Excel metrics','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkPowerPoint      = uicheckbox(pnlOutputs,'Position',[420 0 200 20], ...
                          'Text','PowerPoint','Value',true,'ValueChangedFcn',@syncAllCheckbox);

    % Output folder selection
    uilabel(fig,'Position',[20 220 90 22],'Text','Output folder:');
    txtOutputFolder = uieditfield(fig,'text','Position',[100 220 340 22]);
    uibutton(fig,'push','Position',[450 220 120 22],'Text','Browse...', ...
        'ButtonPushedFcn',@selectOutputFolder);

    % Run button
    uibutton(fig,'push','Position',[240 170 120 30],'Text','Run', ...
        'ButtonPushedFcn',@runAnalysis);

    % Status label
    lblStatus = uilabel(fig,'Position',[20 130 560 22],'Text','Ready', ...
                        'HorizontalAlignment','left');

    % Progress dialog handle (created when running)
    prog = [];

    % ---------------------------------------------------------------------
    % Callback: select input file
    % ---------------------------------------------------------------------
    function selectInputFile(~,~)
        [file,path] = uigetfile('*.xlsx','Select Excel file');
        if isequal(file,0), return; end
        txtInputFile.Value = fullfile(path,file);
    end

    % ---------------------------------------------------------------------
    % Callback: select output folder
    % ---------------------------------------------------------------------
    function selectOutputFolder(~,~)
        folder = uigetdir('','Select output folder');
        if isequal(folder,0), return; end
        txtOutputFolder.Value = folder;
    end

    % ---------------------------------------------------------------------
    % Callback: toggle all outputs
    % ---------------------------------------------------------------------
    function toggleAllOutputs(src,~)
        val = logical(src.Value);
        chkActivityProf.Value  = val;
        chkDailyActivity.Value = val;
        chkLowActivity.Value   = val;
        chkHeatmap.Value       = val;
        chkTempProfiles.Value  = val;
        chkCombined.Value      = val;
        chkLightTracker.Value  = val;
        chkLightDist.Value     = val;
        chkExcelSummary.Value  = val;
        chkExcelMetrics.Value  = val;
        chkPowerPoint.Value    = val;
    end

    % ---------------------------------------------------------------------
    % Keep "All outputs" checkbox in sync
    % ---------------------------------------------------------------------
    function syncAllCheckbox(~,~)
        allSelected = all([ ...
            chkActivityProf.Value, ...
            chkDailyActivity.Value, ...
            chkLowActivity.Value, ...
            chkHeatmap.Value, ...
            chkTempProfiles.Value, ...
            chkCombined.Value, ...
            chkLightTracker.Value, ...
            chkLightDist.Value, ...
            chkExcelSummary.Value, ...
            chkExcelMetrics.Value, ...
            chkPowerPoint.Value ]);
        chkAll.Value = logical(allSelected);
    end

    % ---------------------------------------------------------------------
    % Main analysis function
    % ---------------------------------------------------------------------
    function runAnalysis(~,~)
        try
            % Basic output selection check
            outputs = [ ...
               chkActivityProf.Value, ...
               chkDailyActivity.Value, ...
               chkLowActivity.Value, ...
               chkHeatmap.Value, ...
               chkTempProfiles.Value, ...
               chkCombined.Value, ...
               chkLightTracker.Value, ...
               chkLightDist.Value, ...
               chkExcelSummary.Value, ...
               chkExcelMetrics.Value, ...
               chkPowerPoint.Value ];

            if ~any(outputs)
                uialert(fig,'Please select at least one output.','Output Selection Error');
                return;
            end

            % Create progress dialog for this run
            prog = uiprogressdlg(fig, 'Title','Running analysis', ...
                'Message','Starting...', 'Value',0, 'Cancelable',false);

            setStatus('Validating inputs...', 0.02);

            % Validate input file
            inputFilePath = txtInputFile.Value;
            if isempty(inputFilePath) || ~isfile(inputFilePath)
                uialert(fig,'Please select a valid input file.','Input Error');
                closeProgressDialog();
                return;
            end

            % Validate output folder
            outputFolderPath = txtOutputFolder.Value;
            if isempty(outputFolderPath) || ~isfolder(outputFolderPath)
                uialert(fig,'Please select a valid output folder.','Output Error');
                closeProgressDialog();
                return;
            end

            % --------------------------------------------------------------
            % Load and validate raw data
            % --------------------------------------------------------------
            setStatus('Reading Excel workbook...', 0.05);
            [rawTable, meta, usedSheet] = loadRawData(inputFilePath); %#ok<NASGU>

            setStatus('Parsing timestamps...', 0.10);
            timestamps = parseTimestampsSoft(rawTable.(meta.timeVar));

            validTS = ~isnat(timestamps);
            if nnz(validTS) < 10
                uialert(fig,'Too few valid timestamps after parsing. Check the Time stamp column.','Timestamp Error');
                closeProgressDialog();
                return;
            end

            setStatus('Extracting signals...', 0.12);
            activityCounts = rawTable.(meta.activityVar);
            lightLevels    = rawTable.(meta.lightVar);
            temperature    = rawTable.(meta.tempVar);

            activityCounts = activityCounts(validTS);
            lightLevels    = lightLevels(validTS);
            temperature    = temperature(validTS);
            timestamps     = timestamps(validTS);

            setStatus('Sorting by time...', 0.14);
            [timestamps, sortIdx] = sort(timestamps);
            activityCounts = activityCounts(sortIdx);
            lightLevels    = lightLevels(sortIdx);
            temperature    = temperature(sortIdx);

            % --------------------------------------------------------------
            % Detect sampling interval
            % --------------------------------------------------------------
            lightThreshold = numThreshold.Value;

            setStatus('Detecting sampling interval...', 0.18);
            timeDiffs = diff(timestamps);
            timeDiffs = timeDiffs(timeDiffs > seconds(0));

            if isempty(timeDiffs)
                uialert(fig,'Could not detect sampling interval. Check timestamps.','Sampling Error');
                closeProgressDialog();
                return;
            end

            medianInterval = median(timeDiffs);
            intervalMins   = minutes(medianInterval);

            if ~isfinite(intervalMins) || intervalMins <= 0
                uialert(fig,'Detected sampling interval is invalid. Check timestamps.','Sampling Error');
                closeProgressDialog();
                return;
            end

            binsPerDay = round(1440 / intervalMins);
            if binsPerDay < 1
                uialert(fig,'Detected sampling interval is too large to form daily profiles.','Sampling Error');
                closeProgressDialog();
                return;
            end

            timeAxis = (0:binsPerDay-1) * intervalMins;

            % --------------------------------------------------------------
            % Bin into daily matrices (keeps original strategy)
            % --------------------------------------------------------------
            setStatus('Binning data into daily matrices...', 0.25);

            firstDayStart = dateshift(timestamps(1),'start','day');
            lastDayStart  = dateshift(timestamps(end),'start','day');

            totalDays = days(lastDayStart - firstDayStart) + 1;
            totalDays = max(1, ceil(totalDays));

            elapsedMins = minutes(timestamps - firstDayStart);

            binEdges = 0 : intervalMins : (1440 * totalDays);
            if binEdges(end) < (1440 * totalDays)
                binEdges(end+1) = 1440 * totalDays;
            end

            binIndices = discretize(elapsedMins, binEdges);

            % Handle samples exactly on the final edge
            onFinalEdge = (elapsedMins == binEdges(end));
            if any(onFinalEdge)
                binIndices(onFinalEdge) = numel(binEdges) - 1;
            end

            dayStartDates = firstDayStart + days(0:(totalDays-1))';
            weekDayNames  = cellstr(datestr(dayStartDates,'dddd'));

            binnedActivity    = nan(totalDays, binsPerDay);
            binnedLight       = nan(totalDays, binsPerDay);
            binnedTemperature = nan(totalDays, binsPerDay);

            % Populate matrices
            setStatus('Assigning samples to daily bins...', 0.28);
            for dayIdx = 1:totalDays
                if mod(dayIdx, max(1, round(totalDays/20))) == 0
                    setStatus(sprintf('Assigning bins: day %d/%d...', dayIdx, totalDays), ...
                              0.28 + 0.07*(dayIdx/totalDays));
                end

                low  = (dayIdx-1)*binsPerDay + 1;
                high = dayIdx*binsPerDay;

                sel = binIndices >= low & binIndices <= high & ~isnan(binIndices);
                if ~any(sel)
                    continue;
                end

                relBin = binIndices(sel) - (dayIdx-1)*binsPerDay;
                binnedActivity(dayIdx, relBin)    = activityCounts(sel);
                binnedLight(dayIdx, relBin)       = lightLevels(sel);
                binnedTemperature(dayIdx, relBin) = temperature(sel);
            end

            % --------------------------------------------------------------
            % Complete-day filtering (optional)
            % --------------------------------------------------------------
            if chkCompleteOnly.Value
                setStatus('Filtering to complete 24-hour days only...', 0.36);

                completeMask      = all(~isnan(binnedActivity), 2);

                dayStartDates     = dayStartDates(completeMask);
                weekDayNames      = weekDayNames(completeMask);
                binnedActivity    = binnedActivity(completeMask,:);
                binnedLight       = binnedLight(completeMask,:);
                binnedTemperature = binnedTemperature(completeMask,:);
                totalDays         = sum(completeMask);

                if totalDays == 0
                    uialert(fig,'No complete 24-hour days were found after filtering.','Complete Day Filter');
                    closeProgressDialog();
                    return;
                end
            end

            % --------------------------------------------------------------
            % Compute daily metrics
            % --------------------------------------------------------------
            setStatus('Computing daily summary metrics...', 0.40);

            numberOfWeeks = ceil(totalDays / 7);

            totalActivityPerDay = nansum(binnedActivity,2);
            hoursInLightPerDay  = sum(binnedLight > lightThreshold,2) / (60 / intervalMins);
            minTempPerDay       = min(binnedTemperature,[],2,'omitnan');
            maxTempPerDay       = max(binnedTemperature,[],2,'omitnan');

            % --------------------------------------------------------------
            % L5/M10 (NaN-robust)
            % --------------------------------------------------------------
            setStatus('Computing L5 and M10...', 0.44);

            L5Bins  = max(1, round(5*60/intervalMins));
            M10Bins = max(1, round(10*60/intervalMins));
            minFracValid = 0.90;

            L5Start  = NaT(totalDays,1);
            M10Start = NaT(totalDays,1);
            L5Mean   = nan(totalDays,1);
            M10Mean  = nan(totalDays,1);

            for d = 1:totalDays
                if mod(d, max(1, round(totalDays/20))) == 0
                    setStatus(sprintf('Computing L5/M10: day %d/%d...', d, totalDays), ...
                              0.44 + 0.05*(d/totalDays));
                end

                sig = binnedActivity(d,:);

                cL5  = slidingMeanNan(sig, L5Bins,  minFracValid);
                cM10 = slidingMeanNan(sig, M10Bins, minFracValid);

                if ~all(isnan(cL5))
                    [L5Mean(d), idx5] = min(cL5, [], 'omitnan');
                    L5Start(d) = dayStartDates(d) + minutes((idx5-1) * intervalMins);
                end

                if ~all(isnan(cM10))
                    [M10Mean(d), idx10] = max(cM10, [], 'omitnan');
                    M10Start(d) = dayStartDates(d) + minutes((idx10-1) * intervalMins);
                end
            end

            L5TimeStrings  = strings(totalDays,1);
            M10TimeStrings = strings(totalDays,1);
            for d = 1:totalDays
                if ~isnat(L5Start(d)),  L5TimeStrings(d)  = string(datestr(L5Start(d),  'HH:MM:SS')); end
                if ~isnat(M10Start(d)), M10TimeStrings(d) = string(datestr(M10Start(d), 'HH:MM:SS')); end
            end

            % --------------------------------------------------------------
            % IS/IV hourly
            % --------------------------------------------------------------
            setStatus('Computing IS and IV (hourly)...', 0.52);
            [ISvalue, IVvalue] = computeISIV_hourlyFromDaily(binnedActivity, intervalMins);

            % --------------------------------------------------------------
            % Plotting and exporting
            % --------------------------------------------------------------
            setStatus('Generating figures...', 0.55);

            globalMinTemp = 15;
            globalMaxTemp = 50;
            tempTicks     = [globalMinTemp, globalMinTemp+20, globalMaxTemp];

            % Daily light tracker
            if chkLightTracker.Value
                lightTrackerFolder = fullfile(outputFolderPath,'DailyLightTracker');
                if ~isfolder(lightTrackerFolder), mkdir(lightTrackerFolder); end

                for d = 1:totalDays
                    setStatus(sprintf('Daily light tracker: day %d/%d...', d, totalDays), ...
                              0.55 + 0.15*(d/totalDays));

                    dateLabel = datestr(dayStartDates(d),'dd_mm_yyyy');
                    figLT = figure('Visible','off','Color','w','Position',[100 100 800 600]);

                    ax1 = subplot(2,1,1);
                    plot(ax1, timeAxis, binnedLight(d,:), 'Color',[0 0.5 0.5],'LineWidth',1);
                    hold(ax1,'on');
                    yline(ax1, lightThreshold,':','Color',[0.5 0.5 0.5],'LineWidth',0.5);
                    hold(ax1,'off');
                    title(ax1, sprintf('Light Tracker - %s (Full Range)', datestr(dayStartDates(d),'dd mmm yyyy')), ...
                          'FontSize',14,'HorizontalAlignment','center');
                    xlim(ax1,[0 1440]);
                    ax1.XTick = [];
                    set(ax1,'TickDir','out','Box','off','FontSize',11);
                    ylabel(ax1,'Light (LUX)','FontSize',12);

                    ax2 = subplot(2,1,2);
                    plot(ax2, timeAxis, binnedLight(d,:), 'Color',[0 0.5 0.5],'LineWidth',1);
                    hold(ax2,'on');
                    yline(ax2, lightThreshold,':','Color',[0.5 0.5 0.5],'LineWidth',0.5);
                    hold(ax2,'off');
                    title(ax2,'Light Tracker - 0 to 10 LUX','FontSize',14,'HorizontalAlignment','center');
                    xlim(ax2,[0 1440]);
                    ylim(ax2,[0 10]);
                    set(ax2,'TickDir','out','Box','off','FontSize',11);
                    xlabel(ax2,'Time of Day','FontSize',12);
                    ylabel(ax2,'Light (LUX)','FontSize',12);
                    xticks(ax2,0:240:1440);
                    xticklabels(ax2,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});

                    exportgraphics(figLT, fullfile(lightTrackerFolder, sprintf('DailyLightTracker_%s.jpg',dateLabel)), 'Resolution',600);
                    close(figLT);
                end
            end

            % Weekly outputs
            for wk = 1:numberOfWeeks
                setStatus(sprintf('Weekly figures: week %d/%d...', wk, numberOfWeeks), ...
                          0.70 + 0.20*(wk/numberOfWeeks));

                daysIdx    = (wk-1)*7 + (1:7);
                daysIdx    = daysIdx(daysIdx <= totalDays);
                nThisWeek  = numel(daysIdx);

                dateLabels = cellstr(datestr(dayStartDates(daysIdx),'dd/mm'));
                anyDays    = any(strcmp(ddlAxisStyle.Value,{'Days only','Both'}));
                anyDates   = any(strcmp(ddlAxisStyle.Value,{'Dated only','Both'}));

                % 1) Activity profiles
                if chkActivityProf.Value
                    if anyDays
                        setStatus(sprintf('Week %d/%d: activity profiles (days)...', wk, numberOfWeeks), []);
                        figAP = figure('Visible','off','Color','w');
                        for i = 1:nThisWeek
                            d = daysIdx(i);
                            ax = subplot(nThisWeek,1,i); hold(ax,'on');

                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;
                            mA = binnedActivity(d,:);

                            yMax = max(mA,[],'omitnan');
                            if isnan(yMax) || yMax == 0, yMax = 1; end
                            yMax = yMax * 1.2;

                            fillSegments(ax,mD,timeAxis,0,yMax,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,0,yMax,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,mA,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');

                            ylim(ax,[0 yMax]);
                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);
                            xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            ylabel(ax,sprintf('Day %d',d),'FontSize',12);
                            set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
                            if i==1, title(ax,'Participant Activity Profile','FontSize',16); end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figAP, fullfile(outputFolderPath, sprintf('01_ActivityProfile_Week%d.jpg',wk)),'Resolution',600);
                        close(figAP);
                    end

                    if anyDates
                        setStatus(sprintf('Week %d/%d: activity profiles (dated)...', wk, numberOfWeeks), []);
                        figAPD = figure('Visible','off','Color','w');
                        for i = 1:nThisWeek
                            d = daysIdx(i);
                            ax = subplot(nThisWeek,1,i); hold(ax,'on');

                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;
                            mA = binnedActivity(d,:);

                            yMax = max(mA,[],'omitnan');
                            if isnan(yMax) || yMax == 0, yMax = 1; end
                            yMax = yMax * 1.2;

                            fillSegments(ax,mD,timeAxis,0,yMax,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,0,yMax,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,mA,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');

                            ylim(ax,[0 yMax]);
                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);
                            xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            ylabel(ax,dateLabels{i},'FontSize',12);
                            set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
                            if i==1, title(ax,'Participant Activity Profile (dd/mm)','FontSize',16); end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figAPD, fullfile(outputFolderPath, sprintf('01_ActivityProfile_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figAPD);
                    end
                end

                % 2) Daily activity bar charts
                if chkDailyActivity.Value
                    if anyDays
                        setStatus(sprintf('Week %d/%d: daily activity bars (days)...', wk, numberOfWeeks), []);
                        figDA = figure('Visible','off','Color','w');
                        bar(1:nThisWeek, totalActivityPerDay(daysIdx), 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        xticks(1:nThisWeek);
                        xticklabels(arrayfun(@(x)sprintf('Day %d',x),daysIdx,'UniformOutput',false));
                        xlabel('Day','FontSize',14);
                        title(sprintf('Total Activity by Day (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figDA, fullfile(outputFolderPath, sprintf('02_DailyActivity_Week%d.jpg',wk)),'Resolution',600);
                        close(figDA);
                    end

                    if anyDates
                        setStatus(sprintf('Week %d/%d: daily activity bars (dated)...', wk, numberOfWeeks), []);
                        figDAD = figure('Visible','off','Color','w');
                        bar(1:nThisWeek, totalActivityPerDay(daysIdx), 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        xticks(1:nThisWeek);
                        xticklabels(dateLabels);
                        xlabel('Date','FontSize',14);
                        title(sprintf('Total Activity by Date (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figDAD, fullfile(outputFolderPath, sprintf('02_DailyActivity_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figDAD);
                    end
                end

                % 3) Low-activity call-outs
                if chkLowActivity.Value
                    setStatus(sprintf('Week %d/%d: low-activity call-outs...', wk, numberOfWeeks), []);
                    weekTotals = totalActivityPerDay(daysIdx);
                    lowThresh  = mean(weekTotals,'omitnan') - std(weekTotals,'omitnan');
                    lowIdxs    = find(weekTotals < lowThresh);

                    if anyDays
                        figLC = figure('Visible','off','Color','w');
                        bar(1:nThisWeek, weekTotals, 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        hold on;
                        for idx = lowIdxs'
                            text(idx, weekTotals(idx)+0.05*max(weekTotals),'Low','Color','r', ...
                                'FontSize',12,'HorizontalAlignment','center');
                        end
                        hold off;
                        xticks(1:nThisWeek);
                        xticklabels(arrayfun(@(x)sprintf('Day %d',x),daysIdx,'UniformOutput',false));
                        xlabel('Day','FontSize',14);
                        title(sprintf('Low-Activity Call-outs (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figLC, fullfile(outputFolderPath, sprintf('03_LowActivity_Week%d.jpg',wk)),'Resolution',600);
                        close(figLC);
                    end

                    if anyDates
                        figLCD = figure('Visible','off','Color','w');
                        bar(1:nThisWeek, weekTotals, 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        hold on;
                        for idx = lowIdxs'
                            text(idx, weekTotals(idx)+0.05*max(weekTotals),'Low','Color','r', ...
                                'FontSize',12,'HorizontalAlignment','center');
                        end
                        hold off;
                        xticks(1:nThisWeek);
                        xticklabels(dateLabels);
                        xlabel('Date','FontSize',14);
                        title(sprintf('Low-Activity Call-outs by Date (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figLCD, fullfile(outputFolderPath, sprintf('03_LowActivity_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figLCD);
                    end
                end

                % 4) Activity heatmaps
                if chkHeatmap.Value
                    setStatus(sprintf('Week %d/%d: activity heatmap...', wk, numberOfWeeks), []);
                    blockAct = binnedActivity(daysIdx,:);
                    figHM = figure('Visible','off','Color','w');
                    imagesc(timeAxis,1:nThisWeek,blockAct);
                    axis xy; set(gca,'YDir','reverse');

                    v = blockAct(~isnan(blockAct));
                    if isempty(v)
                        cMin = 0; cMax = 1;
                    else
                        cMin = prctile(v,5);
                        cMax = prctile(v,95);
                        if cMin == cMax, cMax = cMin + 1; end
                    end
                    caxis([cMin cMax]);

                    colormap(parula); colorbar('eastoutside');
                    xlabel('Time of Day','FontSize',12);
                    if anyDates
                        yticks(1:nThisWeek); yticklabels(dateLabels);
                    else
                        yticks(1:nThisWeek); yticklabels(arrayfun(@(x)sprintf('Day %d',x),daysIdx,'UniformOutput',false));
                    end
                    xticks(0:360:1440); xticklabels({'00:00','06:00','12:00','18:00','24:00'});
                    set(gca,'TickDir','out','FontSize',11,'Box','off');
                    title(sprintf('Activity Heatmap (Week %d)',wk),'FontSize',16);
                    exportgraphics(figHM, fullfile(outputFolderPath, sprintf('04_ActivityHeatmap_Week%d.jpg',wk)),'Resolution',600);
                    close(figHM);
                end

                % 5) Temperature profiles
                if chkTempProfiles.Value
                    if anyDays
                        setStatus(sprintf('Week %d/%d: temperature profiles (days)...', wk, numberOfWeeks), []);
                        figTP = figure('Visible','off','Color','w');
                        for i = 1:nThisWeek
                            d = daysIdx(i);
                            ax = subplot(nThisWeek,1,i); hold(ax,'on');

                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;

                            fillSegments(ax,mD,timeAxis,globalMinTemp,globalMaxTemp,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,globalMinTemp,globalMaxTemp,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,binnedTemperature(d,:),'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');

                            ylabel(ax,sprintf('Day %d',d),'FontSize',12);
                            ylim(ax,[globalMinTemp globalMaxTemp]);
                            yticks(ax,tempTicks);
                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);
                            xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            set(ax,'TickDir','out','Box','off','FontSize',11);
                            if i==1, title(ax,sprintf('Temperature Profile - Week %d',wk),'FontSize',16); end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figTP, fullfile(outputFolderPath, sprintf('05_TemperatureProfile_Week%d.jpg',wk)),'Resolution',600);
                        close(figTP);
                    end

                    if anyDates
                        setStatus(sprintf('Week %d/%d: temperature profiles (dated)...', wk, numberOfWeeks), []);
                        figTPD = figure('Visible','off','Color','w');
                        for i = 1:nThisWeek
                            d = daysIdx(i);
                            ax = subplot(nThisWeek,1,i); hold(ax,'on');

                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;

                            fillSegments(ax,mD,timeAxis,globalMinTemp,globalMaxTemp,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,globalMinTemp,globalMaxTemp,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,binnedTemperature(d,:),'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');

                            ylabel(ax,dateLabels{i},'FontSize',12);
                            ylim(ax,[globalMinTemp globalMaxTemp]);
                            yticks(ax,tempTicks);
                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);
                            xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            set(ax,'TickDir','out','Box','off','FontSize',11);
                            if i==1, title(ax,sprintf('Temperature Profile - Week %d (dd/mm)',wk),'FontSize',16); end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figTPD, fullfile(outputFolderPath, sprintf('05_TemperatureProfile_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figTPD);
                    end
                end

                % 6) Weekly light distribution (computed from binnedLight)
                if chkLightDist.Value
                    setStatus(sprintf('Week %d/%d: light distribution...', wk, numberOfWeeks), []);

                    startDateStr = datestr(dayStartDates(daysIdx(1)),'dd/mm');
                    endDateStr   = datestr(dayStartDates(daysIdx(end)),'dd/mm');

                    figLD = figure('Visible','off','Color','w');

                    hourlyAvgWeek = nan(24,1);
                    binHour = floor(timeAxis/60); binHour(binHour > 23) = 23;

                    weekLight = binnedLight(daysIdx,:);
                    for h = 0:23
                        vals = weekLight(:, binHour==h);
                        hourlyAvgWeek(h+1) = mean(vals(:), 'omitnan');
                    end
                    hourlyAvgWeek(isnan(hourlyAvgWeek)) = 0;

                    area(0:23, hourlyAvgWeek, 'FaceColor',[0 0.5 0.5]);
                    title(sprintf('Light Distribution - Week %d (%s to %s)', wk, startDateStr, endDateStr), 'FontSize',16);
                    xlabel('Hour of Day (0 = Midnight)','FontSize',14);
                    ylabel('Average Light (LUX)','FontSize',14);
                    xlim([0 23]); xticks(0:1:23);
                    set(gca,'TickDir','out','FontSize',11,'Box','off');

                    exportgraphics(figLD, fullfile(outputFolderPath, sprintf('06_LightDistribution_Week%d.jpg',wk)),'Resolution',600);
                    close(figLD);
                end

                % 7) Combined profile
                if chkCombined.Value
                    setStatus(sprintf('Week %d/%d: combined profile...', wk, numberOfWeeks), []);
                    figCP = figure('Visible','off','Color','w');

                    for i = 1:nThisWeek
                        d = daysIdx(i);

                        axMain = subplot(nThisWeek,1,i);
                        hold(axMain,'on');
                        axMain.XAxisLocation = 'origin';
                        axMain.XLim       = [0 1440];
                        axMain.XTick      = 0:240:1440;
                        axMain.XTickLabel = {'00:00','04:00','08:00','12:00','16:00','20:00','24:00'};

                        mL = binnedLight(d,:) > lightThreshold;
                        mD = ~mL;

                        yMax = max(binnedActivity(d,:),[],'omitnan');
                        if isnan(yMax) || yMax == 0, yMax = 1; end
                        yMax = yMax * 1.2;

                        fillSegments(axMain,mD,timeAxis,0,yMax,[0.9 0.9 0.9]);
                        fillSegments(axMain,mL,timeAxis,0,yMax,[0.9290 0.6940 0.1250]);
                        bar(axMain,timeAxis,binnedActivity(d,:),'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');

                        ylim(axMain,[0 yMax]);
                        axMain.YTick = [0 yMax];
                        set(axMain,'YTickLabel',[],'Box','off','TickDir','out','FontSize',11);

                        if anyDates
                            ylabel(axMain,dateLabels{i},'FontSize',12);
                        else
                            ylabel(axMain,sprintf('Day %d',d),'FontSize',12);
                        end

                        pos    = axMain.Position;
                        axTemp = axes('Position',pos, 'Color','none', 'YAxisLocation','right', ...
                                      'XAxisLocation','bottom', 'Box','off', 'TickDir','out', 'FontSize',11, ...
                                      'XLim',axMain.XLim, 'XTick',axMain.XTick);
                        hold(axTemp,'on');
                        axTemp.XTick = [];
                        axTemp.XAxis.Visible = 'off';

                        plot(axTemp,timeAxis,binnedTemperature(d,:),'LineWidth',1.5,'Color',[1 0 0]);
                        ylim(axTemp,[0 globalMaxTemp]);
                        yticks(axTemp,[0 tempTicks]);
                        ylabel(axTemp,'°C','FontSize',12);

                        linkaxes([axMain,axTemp],'x');

                        if i==1, title(axMain,sprintf('Combined Profile - Week %d',wk),'FontSize',16); end

                        hold(axMain,'off');
                        hold(axTemp,'off');
                    end

                    xlabel(axMain,'Time of Day','FontSize',12);
                    exportgraphics(figCP, fullfile(outputFolderPath, sprintf('07_CombinedProfile_Week%d.jpg',wk)),'Resolution',600);
                    close(figCP);
                end
            end

            % --------------------------------------------------------------
            % Write Excel workbook
            % --------------------------------------------------------------
            setStatus('Writing Excel outputs...', 0.92);

            summaryFile = fullfile(outputFolderPath,'10_Participant_Results.xlsx');

            if chkExcelSummary.Value
                setStatus('Writing Excel: Summary sheet...', 0.93);
                SummaryTable = table( ...
                    dayStartDates, ...
                    weekDayNames, ...
                    (1:totalDays)', ...
                    totalActivityPerDay, ...
                    hoursInLightPerDay, ...
                    L5TimeStrings, ...
                    L5Mean, ...
                    M10TimeStrings, ...
                    M10Mean, ...
                    minTempPerDay, ...
                    maxTempPerDay, ...
                    'VariableNames',{ ...
                      'Date','Weekday','Day','TotalActivity','HoursInLight', ...
                      'L5_StartTime','L5_Mean','M10_StartTime','M10_Mean', ...
                      'MinTemperature','MaxTemperature'} );
                writetable(SummaryTable, summaryFile, 'Sheet','Summary');
            end

            if chkExcelMetrics.Value
                setStatus('Writing Excel: Metrics sheet...', 0.95);
                MetricsTable = table( ...
                    intervalMins, ...
                    lightThreshold, ...
                    logical(chkCompleteOnly.Value), ...
                    ISvalue, ...
                    IVvalue, ...
                    'VariableNames',{ ...
                        'SamplingInterval_Minutes', ...
                        'LightThreshold_Lux', ...
                        'CompleteDaysOnly', ...
                        'InterdailyStability_Hourly', ...
                        'IntradailyVariability_Hourly'} );
                writetable(MetricsTable, summaryFile, 'Sheet','Metrics');
            end

            setStatus('Writing Excel: Definitions sheet...', 0.96);
            definitions = {
              'Date','Calendar date (day start at midnight)','Aligns metrics to calendar days';
              'Weekday','Day of the week','Useful for weekday vs weekend patterns';
              'Day','Sequential day index','Day number since the first analysed day';
              'TotalActivity','Sum of binned activity values across the day','Overall daily movement proxy';
              'HoursInLight',sprintf('Total hours with Light > %.2f lux', lightThreshold),'Duration of light exposure above the threshold';
              'L5_StartTime','Clock time when the five-hour lowest-activity window begins','Rest-activity trough onset estimate';
              'L5_Mean','Mean activity during the L5 window','Depth of rest trough estimate';
              'M10_StartTime','Clock time when the ten-hour highest-activity window begins','Main activity window onset estimate';
              'M10_Mean','Mean activity during the M10 window','Strength of activity peak estimate';
              'MinTemperature','Minimum recorded temperature for the day','Lower bound across available samples';
              'MaxTemperature','Maximum recorded temperature for the day','Upper bound across available samples';
              'InterdailyStability_Hourly','IS computed on hourly mean activity','Higher values indicate more day-to-day regularity';
              'IntradailyVariability_Hourly','IV computed on hourly mean activity','Higher values indicate more within-day fragmentation'
            };
            writecell([{'Term','Definition','Interpretation'}; definitions], summaryFile, 'Sheet','Definitions');

            % --------------------------------------------------------------
            % Compile PowerPoint (includes subfolders)
            % --------------------------------------------------------------
            if chkPowerPoint.Value
                setStatus('Creating PowerPoint...', 0.97);

                if ~exist('mlreportgen.ppt.Presentation','class')
                    uialert(fig,'Report Generator toolbox not found. PowerPoint export cannot run.','PowerPoint Error');
                else
                    import mlreportgen.ppt.*;

                    pptFile = fullfile(outputFolderPath,'AllFigures_Report.pptx');
                    ppt     = Presentation(pptFile);
                    open(ppt);

                    setStatus('PowerPoint: adding run settings slide...', 0.975);
                    slide0 = add(ppt,'Title and Content');
                    replace(slide0,'Title','Run Settings');
                    settingsText = sprintf([ ...
                        'Input file: %s\n' ...
                        'Output folder: %s\n' ...
                        'Sampling interval (minutes): %.4g\n' ...
                        'Light threshold (lux): %.4g\n' ...
                        'Complete days only: %s\n' ...
                        'Axis style: %s\n'], ...
                        inputFilePath, outputFolderPath, intervalMins, lightThreshold, ...
                        mat2str(logical(chkCompleteOnly.Value)), ddlAxisStyle.Value);
                    replace(slide0,'Content', settingsText);

                    setStatus('PowerPoint: collecting JPGs (including subfolders)...', 0.98);
                    jpgFiles = dir(fullfile(outputFolderPath,'**','*.jpg'));

                    if ~isempty(jpgFiles)
                        fullPaths = fullfile({jpgFiles.folder},{jpgFiles.name});
                        [~, order] = sort(lower(fullPaths));
                        jpgFiles = jpgFiles(order);
                    end

                    for k = 1:numel(jpgFiles)
                        setStatus(sprintf('PowerPoint: slide %d/%d...', k, numel(jpgFiles)), ...
                                  0.98 + 0.02*(k/max(1,numel(jpgFiles))));

                        slide = add(ppt,'Title and Content');
                        relPath = erase(fullfile(jpgFiles(k).folder, jpgFiles(k).name), [outputFolderPath filesep]);
                        replace(slide,'Title', erase(relPath, '.jpg'));
                        replace(slide,'Content', Picture(fullfile(jpgFiles(k).folder, jpgFiles(k).name)));
                    end

                    close(ppt);
                end
            end

            % Done
            setStatus('Done.', 1.0);
            closeProgressDialog();
            pause(0.5);
            close(fig);

        catch ME
            closeProgressDialog();
            uialert(fig, ME.message, 'Error');
            lblStatus.Text = 'Error encountered.';
        end
    end

    % ---------------------------------------------------------------------
    % Status and progress helpers
    % ---------------------------------------------------------------------
    function setStatus(msg, frac)
        lblStatus.Text = char(msg);

        if ~isempty(prog) && isvalid(prog)
            prog.Message = char(msg);
            if nargin >= 2 && ~isempty(frac) && isfinite(frac)
                prog.Value = max(0, min(1, frac));
            end
        end

        drawnow limitrate;
    end

    function closeProgressDialog()
        if ~isempty(prog) && isvalid(prog)
            close(prog);
        end
        prog = [];
    end

    % ---------------------------------------------------------------------
    % Helper: fill contiguous segments of a logical mask
    % ---------------------------------------------------------------------
    function fillSegments(ax, maskArray, xVals, yBottom, yTop, fillColor)
        maskArray = logical(maskArray);
        diffMask  = diff([0 maskArray 0]);
        runStarts = find(diffMask==1);
        runEnds   = find(diffMask==-1)-1;

        for r = 1:numel(runStarts)
            idx = runStarts(r):runEnds(r);
            if isempty(idx), continue; end
            xPoly = [xVals(idx), fliplr(xVals(idx))];
            yPoly = [yBottom*ones(1,numel(idx)), yTop*ones(1,numel(idx))];
            fill(ax, xPoly, yPoly, fillColor, 'EdgeColor','none', 'FaceAlpha',0.3);
        end
    end

    % ---------------------------------------------------------------------
    % Helper: load raw data, prefer RawData but fall back safely
    % ---------------------------------------------------------------------
    function [T, meta, usedSheet] = loadRawData(xlsxPath)
        sheets = sheetnames(xlsxPath);

        usedSheet = '';
        if any(strcmpi(sheets,'RawData'))
            usedSheet = sheets{find(strcmpi(sheets,'RawData'),1)};
        else
            usedSheet = sheets{1};
        end

        T = readtable(xlsxPath, 'Sheet', usedSheet, 'VariableNamingRule','preserve');

        meta.timeVar     = requireVarLoose(T, {'Time stamp','Timestamp','Time Stamp','Time'});
        meta.activityVar = requireVarLoose(T, {'Sum of vector (SVMg)','SVMg','SVM','Activity','Activity (SVMg)'});
        meta.lightVar    = requireVarLoose(T, {'Light level (LUX)','Light level (Lux)','Light (LUX)','Lux','Light'});
        meta.tempVar     = requireVarLoose(T, {'Temperature','Temp','Temperature (C)','Temperature (°C)'});
    end

    % ---------------------------------------------------------------------
    % Helper: match variable names ignoring whitespace and case
    % ---------------------------------------------------------------------
    function v = requireVarLoose(T, candidates)
        vars = string(T.Properties.VariableNames);
        normVars = normaliseNames(vars);
        normCands = normaliseNames(string(candidates));

        v = "";
        for i = 1:numel(normCands)
            hit = find(normVars == normCands(i), 1);
            if ~isempty(hit)
                v = vars(hit);
                return;
            end
        end

        error('Missing required column. Tried: %s. Found columns: %s', ...
            strjoin(string(candidates), ', '), strjoin(vars, ', '));
    end

    function n = normaliseNames(s)
        s = lower(string(s));
        s = strip(s);
        s = regexprep(s, '\s+', ' ');
        n = s;
    end

    % ---------------------------------------------------------------------
    % Helper: timestamp parsing that drops failures rather than hard-stopping
    % ---------------------------------------------------------------------
    function ts = parseTimestampsSoft(tsRaw)
        if isdatetime(tsRaw)
            ts = tsRaw;
            return;
        end

        if isnumeric(tsRaw)
            try
                ts = datetime(tsRaw, 'ConvertFrom','excel');
                return;
            catch
                tsStr = string(tsRaw);
            end
        else
            tsStr = string(tsRaw);
        end

        tsStr = strip(tsStr);

        fmts = { ...
            'yyyy-MM-dd HH:mm:ss:SSS', ...
            'yyyy-MM-dd HH:mm:ss.SSS', ...
            'yyyy-MM-dd HH:mm:ss', ...
            'dd/MM/yyyy HH:mm:ss', ...
            'dd/MM/yyyy HH:mm', ...
            'MM/dd/yyyy HH:mm:ss', ...
            'MM/dd/yyyy HH:mm'};

        ts = NaT(size(tsStr));
        for i = 1:numel(fmts)
            try
                tTry = datetime(tsStr, 'InputFormat', fmts{i});
                ok = ~isnat(tTry) & isnat(ts);
                ts(ok) = tTry(ok);
            catch
            end
            if all(~isnat(ts)), break; end
        end
    end

    % ---------------------------------------------------------------------
    % Helper: NaN-robust sliding mean with minimum valid fraction
    % ---------------------------------------------------------------------
    function m = slidingMeanNan(x, win, minFrac)
        x = double(x(:))'; % row
        n = numel(x);

        if win > n
            m = nan(1, max(1, n-win+1));
            return;
        end

        valid = ~isnan(x);
        x0 = x;
        x0(~valid) = 0;

        kernel = ones(1,win);
        sumX   = conv(x0, kernel, 'valid');
        cntX   = conv(double(valid), kernel, 'valid');

        minCnt = ceil(minFrac * win);
        m = sumX ./ cntX;
        m(cntX < minCnt) = nan;
    end

    % ---------------------------------------------------------------------
    % Helper: compute IS/IV on hourly means derived from daily matrix
    % ---------------------------------------------------------------------
    function [ISvalue, IVvalue] = computeISIV_hourlyFromDaily(dailyMat, intervalMins)
        [nDays, nBins] = size(dailyMat);

        timeAxisLocal = (0:nBins-1) * intervalMins;
        hourIdx = floor(timeAxisLocal/60) + 1;
        hourIdx(hourIdx < 1) = 1;
        hourIdx(hourIdx > 24) = 24;

        hourlyMat = nan(nDays, 24);
        for h = 1:24
            cols = (hourIdx == h);
            if any(cols)
                hourlyMat(:,h) = mean(dailyMat(:,cols), 2, 'omitnan');
            end
        end

        x = hourlyMat(:);
        valid = ~isnan(x);
        if nnz(valid) < 48
            ISvalue = nan;
            IVvalue = nan;
            return;
        end

        xValid = x(valid);
        grandMean = mean(xValid,'omitnan');

        meanByHour = mean(hourlyMat, 1, 'omitnan'); % 1x24
        p = 24;

        denom = sum((xValid - grandMean).^2, 'omitnan');
        numer = p * sum((meanByHour - grandMean).^2, 'omitnan');

        if denom <= 0 || isnan(denom)
            ISvalue = nan;
        else
            ISvalue = numer / denom;
        end

        xGrid = x; % includes NaNs
        validPairs = ~isnan(xGrid(1:end-1)) & ~isnan(xGrid(2:end));
        if nnz(validPairs) < 24
            IVvalue = nan;
            return;
        end

        dx = diff(xGrid);
        mssd = mean(dx(validPairs).^2, 'omitnan');
        varx = mean((xValid - grandMean).^2, 'omitnan');

        if varx <= 0 || isnan(varx)
            IVvalue = nan;
        else
            IVvalue = mssd / varx;
        end
    end

end