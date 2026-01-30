function Actograph_v7_forparticipants_GUI
% -------------------------------------------------------------------------
% Actograph Analysis and Summary GUI for GENEActiv Participants
% -------------------------------------------------------------------------
% A self-contained MATLAB GUI for processing wrist-worn actigraphy data
% exported from GENEActiv and placed into an Excel workbook.
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
%   - High-resolution JPGs (paged by "Days per figure", or "All")
%   - Daily light tracker JPGs (one per day, in subfolder)
%       * 3 panels: (1) full range (2) hourly mean ± SD (3) 0–10 lux
%   - Light distribution plots per part (or All)
%       * Area plot of hourly mean
%       * Line plot with mean ± SD (lower band clipped at 0 lux)
%   - Activity heatmaps per part (or All)
%       * Colourbar labelled "Activity Intensity"
%       * Y labels: Day, Date, or Day + Date (Both)
%       * X axis labelled 00:00 ... 24:00 (circular completion column added)
%   - Activity and temperature profiles
%       * Row labels are horizontal and moved left to avoid overlap
%       * Temperature y-ticks simplify to [15 50] when many days shown
%   - Excel workbook with Summary, Metrics, Definitions
%   - Optional PowerPoint compiling all JPGs (including subfolders)
%       * Uses "Title and Content" layout
%       * No settings slide (removed for robustness)
%       * Deterministic ordering by sorted full path (string sort)
%       * Slide titles derived from folder (if any) + filename
%
% Key behaviour
%   - No week separation: outputs are generated in sequential blocks determined
%     by "Days per figure" (or one block if "All" is selected)
%   - Header matching ignores extra whitespace (handles e.g., "Temperature ")
%   - Timestamp parsing tries multiple formats and drops invalid rows
%   - L5/M10 computed using NaN-robust sliding means (requires enough valid bins)
%   - IS/IV computed on hourly means derived from the binned daily matrix
%   - Status label + progress dialog report what the code is doing throughout
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
    fig = uifigure('Name','Actograph Analysis','Position',[300 200 600 620]);

    % Input file selection
    uilabel(fig,'Position',[20 580 70 22],'Text','Input file:');
    txtInputFile = uieditfield(fig,'text','Position',[100 580 340 22]);
    uibutton(fig,'push','Position',[450 580 120 22],'Text','Browse...', ...
        'ButtonPushedFcn',@selectInputFile);

    % Global light threshold
    uilabel(fig,'Position',[20 540 140 22],'Text','Light threshold (lux):');
    numThreshold = uieditfield(fig,'numeric','Position',[160 540 80 22], ...
        'Value',5,'Limits',[0 Inf],'RoundFractionalValues',true);

    % Complete days only checkbox
    chkCompleteOnly = uicheckbox(fig,'Position',[20 500 280 22], ...
        'Text','Include only complete 24-hour days');

    % Axis style dropdown
    uilabel(fig,'Position',[20 460 80 22],'Text','Axis style:');
    ddlAxisStyle = uidropdown(fig,'Position',[100 460 200 22], ...
        'Items',{'Days only','Dated only','Both'},'Value','Both');

    % Days per figure (paging) controls
    uilabel(fig,'Position',[320 460 110 22],'Text','Days per figure:');
    numDaysPerFigure = uieditfield(fig,'numeric','Position',[430 460 60 22], ...
        'Value',14,'Limits',[1 Inf],'RoundFractionalValues',true);
    chkAllDaysPerFigure = uicheckbox(fig,'Position',[500 460 80 22], ...
        'Text','All','Value',false,'ValueChangedFcn',@toggleDaysPerFigureMode);

    % Outputs selection panel
    pnlOutputs = uipanel(fig,'Title','Select Outputs','Position',[20 300 560 140]);

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
    uilabel(fig,'Position',[20 260 90 22],'Text','Output folder:');
    txtOutputFolder = uieditfield(fig,'text','Position',[100 260 340 22]);
    uibutton(fig,'push','Position',[450 260 120 22],'Text','Browse...', ...
        'ButtonPushedFcn',@selectOutputFolder);

    % Run button
    uibutton(fig,'push','Position',[240 210 120 30],'Text','Run', ...
        'ButtonPushedFcn',@runAnalysis);

    % Status label
    lblStatus = uilabel(fig,'Position',[20 170 560 22],'Text','Ready', ...
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
    % Callback: toggle "All" for days-per-figure
    % ---------------------------------------------------------------------
    function toggleDaysPerFigureMode(src,~)
        numDaysPerFigure.Enable = ternary(~logical(src.Value), 'on', 'off');
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

            prog = uiprogressdlg(fig, 'Title','Running analysis', ...
                'Message','Starting...', 'Value',0, 'Cancelable',false);

            setStatus('Validating inputs...', 0.02);

            inputFilePath = txtInputFile.Value;
            if isempty(inputFilePath) || ~isfile(inputFilePath)
                uialert(fig,'Please select a valid input file.','Input Error');
                closeProgressDialog();
                return;
            end

            outputFolderPath = txtOutputFolder.Value;
            if isempty(outputFolderPath) || ~isfolder(outputFolderPath)
                uialert(fig,'Please select a valid output folder.','Output Error');
                closeProgressDialog();
                return;
            end

            if chkAllDaysPerFigure.Value
                daysPerFigure = 0; % 0 means "All"
            else
                daysPerFigure = max(1, round(numDaysPerFigure.Value));
            end

            setStatus('Reading Excel workbook...', 0.05);
            [rawTable, meta, usedSheet] = loadRawData(inputFilePath);

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

            setStatus(sprintf('Sorting by time (sheet: %s)...', usedSheet), 0.14);
            [timestamps, sortIdx] = sort(timestamps);
            activityCounts = activityCounts(sortIdx);
            lightLevels    = lightLevels(sortIdx);
            temperature    = temperature(sortIdx);

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

            onFinalEdge = (elapsedMins == binEdges(end));
            if any(onFinalEdge)
                binIndices(onFinalEdge) = numel(binEdges) - 1;
            end

            dayStartDates = firstDayStart + days(0:(totalDays-1))';
            weekDayNames  = cellstr(datestr(dayStartDates,'dddd'));

            binnedActivity    = nan(totalDays, binsPerDay);
            binnedLight       = nan(totalDays, binsPerDay);
            binnedTemperature = nan(totalDays, binsPerDay);

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

            setStatus('Computing daily summary metrics...', 0.40);

            totalActivityPerDay = nansum(binnedActivity,2);
            hoursInLightPerDay  = sum(binnedLight > lightThreshold,2) / (60 / intervalMins);
            minTempPerDay       = min(binnedTemperature,[],2,'omitnan');
            maxTempPerDay       = max(binnedTemperature,[],2,'omitnan');

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

            setStatus('Computing IS and IV (hourly)...', 0.52);
            [ISvalue, IVvalue] = computeISIV_hourlyFromDaily(binnedActivity, intervalMins);

            setStatus('Generating figures...', 0.55);

            globalMinTemp = 15;
            globalMaxTemp = 50;

            % ------------------------------------------------------------------
            % Daily light tracker (day-by-day, 3 panels)
            % ------------------------------------------------------------------
            if chkLightTracker.Value
                setStatus('Preparing daily light tracker outputs...', 0.56);
                lightTrackerFolder = fullfile(outputFolderPath,'DailyLightTracker');
                if ~isfolder(lightTrackerFolder), mkdir(lightTrackerFolder); end

                for d = 1:totalDays
                    setStatus(sprintf('Daily light tracker: day %d/%d...', d, totalDays), ...
                              0.56 + 0.10*(d/totalDays));

                    dateLabel = datestr(dayStartDates(d),'dd_mm_yyyy');

                    figLT = figure('Visible','off','Color','w','Position',[100 100 900 800]);

                    % Panel 1: full range
                    ax1 = subplot(3,1,1);
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

                    % Panel 2: hourly mean ± SD for that day (lower clipped at 0)
                    axM = subplot(3,1,2); hold(axM,'on');
                    [hourMean, hourSD] = hourlyMeanSD_fromSingleDay(binnedLight(d,:), timeAxis);
                    xH = 0:23;

                    sdLower = max(0, hourMean - hourSD);
                    sdUpper = hourMean + hourSD;

                    bandX = [xH, fliplr(xH)];
                    bandY = [sdLower; flipud(sdUpper)]';
                    fill(axM, bandX, bandY, [0 0.5 0.5], 'FaceAlpha',0.2, 'EdgeColor','none');
                    plot(axM, xH, hourMean, 'LineWidth',2, 'Color',[0 0.5 0.5]);
                    yline(axM, lightThreshold,':','Color',[0.5 0.5 0.5],'LineWidth',0.5);

                    title(axM, 'Hourly Mean ± SD (within day)', 'FontSize',14,'HorizontalAlignment','center');
                    xlim(axM,[0 23]);
                    xticks(axM,0:1:23);
                    xlabel(axM,'Hour of Day','FontSize',12);
                    ylabel(axM,'Light (LUX)','FontSize',12);
                    set(axM,'TickDir','out','Box','off','FontSize',11);
                    hold(axM,'off');

                    % Panel 3: zoom 0–10
                    ax2 = subplot(3,1,3);
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

            % ------------------------------------------------------------------
            % Build block indices (parts)
            % ------------------------------------------------------------------
            if daysPerFigure == 0
                partStarts = 1;
            else
                partStarts = 1:daysPerFigure:totalDays;
            end
            nParts = numel(partStarts);

            anyDaysGlobal  = any(strcmp(ddlAxisStyle.Value,{'Days only','Both'}));
            anyDatesGlobal = any(strcmp(ddlAxisStyle.Value,{'Dated only','Both'}));
            suppressPartTag = (daysPerFigure == 0);

            for p = 1:nParts
                if daysPerFigure == 0
                    daysIdx = 1:totalDays;
                else
                    daysIdx = partStarts(p):min(totalDays, partStarts(p)+daysPerFigure-1);
                end

                nThisBlock  = numel(daysIdx);
                dateLabels  = cellstr(datestr(dayStartDates(daysIdx),'dd/mm'));

                if suppressPartTag
                    blockTag = 'AllDays';
                else
                    blockTag = sprintf('Part%02d', p);
                end

                blockStartStr = datestr(dayStartDates(daysIdx(1)),'yyyy-mm-dd');
                blockEndStr   = datestr(dayStartDates(daysIdx(end)),'yyyy-mm-dd');
                blockRangeStr = sprintf('%s_to_%s', blockStartStr, blockEndStr);

                setStatus(sprintf('Figures: %s (%d days)...', blockTag, nThisBlock), ...
                          0.66 + 0.22*(p/max(1,nParts)));

                % Solution 1: simplify temperature ticks when many days shown
                if nThisBlock > 10
                    tempTicks = [globalMinTemp, globalMaxTemp];
                    tempTickFont = 10;
                else
                    tempTicks = [globalMinTemp, globalMinTemp+20, globalMaxTemp];
                    tempTickFont = 11;
                end

                % Row label placement: option A
                rowLabelX = -0.08;
                leftMarginFrac = 0.08;

                % ----------------------------------------------------------
                % Activity profiles
                % ----------------------------------------------------------
                if chkActivityProf.Value
                    if anyDaysGlobal
                        figAP = figure('Visible','off','Color','w');
                        set(figAP,'Position',[100 100 1200 600]);

                        for i = 1:nThisBlock
                            d = daysIdx(i);
                            ax = subplot(nThisBlock,1,i); hold(ax,'on');
                            shiftAxisRight(ax, leftMarginFrac);

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

                            if i ~= nThisBlock
                                ax.XTickLabel = [];
                            else
                                xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            end

                            addRowLabel(ax, sprintf('Day %d', d), rowLabelX);
                            set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);

                            if i==1
                                title(ax, sprintf('Participant Activity Profile (%s)', strrep(blockRangeStr,'_',' ')), 'FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);

                        outName = sprintf('ActivityProfile_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figAP, fullfile(outputFolderPath, outName), 'Resolution',600);
                        close(figAP);
                    end

                    if anyDatesGlobal
                        figAPD = figure('Visible','off','Color','w');
                        set(figAPD,'Position',[100 100 1200 600]);

                        for i = 1:nThisBlock
                            d = daysIdx(i);
                            ax = subplot(nThisBlock,1,i); hold(ax,'on');
                            shiftAxisRight(ax, leftMarginFrac);

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

                            if i ~= nThisBlock
                                ax.XTickLabel = [];
                            else
                                xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            end

                            addRowLabel(ax, dateLabels{i}, rowLabelX);
                            set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);

                            if i==1
                                title(ax, sprintf('Participant Activity Profile (dd/mm) (%s)', strrep(blockRangeStr,'_',' ')), 'FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);

                        outName = sprintf('ActivityProfile_Dated_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figAPD, fullfile(outputFolderPath, outName), 'Resolution',600);
                        close(figAPD);
                    end
                end

                % ----------------------------------------------------------
                % Daily activity bar charts
                % ----------------------------------------------------------
                if chkDailyActivity.Value
                    if anyDaysGlobal
                        figDA = figure('Visible','off','Color','w');
                        bar(1:nThisBlock, totalActivityPerDay(daysIdx), 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        xticks(1:nThisBlock);
                        xticklabels(arrayfun(@(x)sprintf('Day %d',x),daysIdx,'UniformOutput',false));
                        xlabel('Day','FontSize',14);
                        title(sprintf('Total Activity by Day (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');

                        outName = sprintf('DailyActivity_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figDA, fullfile(outputFolderPath, outName),'Resolution',600);
                        close(figDA);
                    end

                    if anyDatesGlobal
                        figDAD = figure('Visible','off','Color','w');
                        bar(1:nThisBlock, totalActivityPerDay(daysIdx), 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        xticks(1:nThisBlock);
                        xticklabels(dateLabels);
                        xlabel('Date','FontSize',14);
                        title(sprintf('Total Activity by Date (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');

                        outName = sprintf('DailyActivity_Dated_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figDAD, fullfile(outputFolderPath, outName),'Resolution',600);
                        close(figDAD);
                    end
                end

                % ----------------------------------------------------------
                % Low-activity call-outs
                % ----------------------------------------------------------
                if chkLowActivity.Value
                    blockTotals = totalActivityPerDay(daysIdx);
                    lowThresh   = mean(blockTotals,'omitnan') - std(blockTotals,'omitnan');
                    lowIdxs     = find(blockTotals < lowThresh);

                    if anyDaysGlobal
                        figLC = figure('Visible','off','Color','w');
                        bar(1:nThisBlock, blockTotals, 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        hold on;
                        for idx = lowIdxs'
                            text(idx, blockTotals(idx)+0.05*max(blockTotals),'Low','Color','r', ...
                                'FontSize',12,'HorizontalAlignment','center');
                        end
                        hold off;
                        xticks(1:nThisBlock);
                        xticklabels(arrayfun(@(x)sprintf('Day %d',x),daysIdx,'UniformOutput',false));
                        xlabel('Day','FontSize',14);
                        title(sprintf('Low-Activity Call-outs (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');

                        outName = sprintf('LowActivity_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figLC, fullfile(outputFolderPath, outName),'Resolution',600);
                        close(figLC);
                    end

                    if anyDatesGlobal
                        figLCD = figure('Visible','off','Color','w');
                        bar(1:nThisBlock, blockTotals, 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        hold on;
                        for idx = lowIdxs'
                            text(idx, blockTotals(idx)+0.05*max(blockTotals),'Low','Color','r', ...
                                'FontSize',12,'HorizontalAlignment','center');
                        end
                        hold off;
                        xticks(1:nThisBlock);
                        xticklabels(dateLabels);
                        xlabel('Date','FontSize',14);
                        title(sprintf('Low-Activity Call-outs by Date (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');

                        outName = sprintf('LowActivity_Dated_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figLCD, fullfile(outputFolderPath, outName),'Resolution',600);
                        close(figLCD);
                    end
                end

                % ----------------------------------------------------------
                % Activity heatmap
                % ----------------------------------------------------------
                if chkHeatmap.Value
                    blockAct = binnedActivity(daysIdx,:);

                    blockActHM = [blockAct, blockAct(:,1)];
                    timeAxisHM = [timeAxis, 1440];

                    figHM = figure('Visible','off','Color','w');
                    imagesc(timeAxisHM,1:nThisBlock,blockActHM);
                    axis xy; set(gca,'YDir','reverse');

                    v = blockActHM(~isnan(blockActHM));
                    if isempty(v)
                        cMin = 0; cMax = 1;
                    else
                        cMin = prctile(v,5);
                        cMax = prctile(v,95);
                        if cMin == cMax, cMax = cMin + 1; end
                    end
                    caxis([cMin cMax]);

                    colormap(parula);
                    cb = colorbar('eastoutside');
                    cb.Label.String = 'Activity Intensity';
                    cb.Label.FontSize = 12;

                    xlabel('Time of Day','FontSize',12);

                    switch ddlAxisStyle.Value
                        case 'Days only'
                            yLabs = arrayfun(@(x)sprintf('Day %d',x), daysIdx, 'UniformOutput',false);
                        case 'Dated only'
                            yLabs = dateLabels;
                        otherwise
                            yLabs = cell(nThisBlock,1);
                            for ii = 1:nThisBlock
                                yLabs{ii} = sprintf('Day %d (%s)', daysIdx(ii), dateLabels{ii});
                            end
                    end
                    yticks(1:nThisBlock);
                    yticklabels(yLabs);

                    xticks(0:360:1440);
                    xticklabels({'00:00','06:00','12:00','18:00','24:00'});
                    set(gca,'TickDir','out','FontSize',11,'Box','off');
                    title(sprintf('Activity Heatmap (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);

                    outName = sprintf('ActivityHeatmap_%s_%s.jpg', blockTag, blockRangeStr);
                    exportgraphics(figHM, fullfile(outputFolderPath, outName),'Resolution',600);
                    close(figHM);
                end

                % ----------------------------------------------------------
                % Temperature profiles
                % ----------------------------------------------------------
                if chkTempProfiles.Value
                    if anyDaysGlobal
                        figTP = figure('Visible','off','Color','w');
                        set(figTP,'Position',[100 100 1200 600]);

                        for i = 1:nThisBlock
                            d = daysIdx(i);
                            ax = subplot(nThisBlock,1,i); hold(ax,'on');
                            shiftAxisRight(ax, leftMarginFrac);

                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;

                            fillSegments(ax,mD,timeAxis,globalMinTemp,globalMaxTemp,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,globalMinTemp,globalMaxTemp,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,binnedTemperature(d,:),'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');

                            ylim(ax,[globalMinTemp globalMaxTemp]);
                            yticks(ax,tempTicks);
                            set(ax,'FontSize',tempTickFont);

                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);

                            if i ~= nThisBlock
                                ax.XTickLabel = [];
                            else
                                xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            end

                            addRowLabel(ax, sprintf('Day %d', d), rowLabelX);
                            set(ax,'TickDir','out','Box','off');

                            if i==1
                                title(ax, sprintf('Temperature Profile (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);

                        outName = sprintf('TemperatureProfile_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figTP, fullfile(outputFolderPath, outName),'Resolution',600);
                        close(figTP);
                    end

                    if anyDatesGlobal
                        figTPD = figure('Visible','off','Color','w');
                        set(figTPD,'Position',[100 100 1200 600]);

                        for i = 1:nThisBlock
                            d = daysIdx(i);
                            ax = subplot(nThisBlock,1,i); hold(ax,'on');
                            shiftAxisRight(ax, leftMarginFrac);

                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;

                            fillSegments(ax,mD,timeAxis,globalMinTemp,globalMaxTemp,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,globalMinTemp,globalMaxTemp,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,binnedTemperature(d,:),'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');

                            ylim(ax,[globalMinTemp globalMaxTemp]);
                            yticks(ax,tempTicks);
                            set(ax,'FontSize',tempTickFont);

                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);

                            if i ~= nThisBlock
                                ax.XTickLabel = [];
                            else
                                xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            end

                            addRowLabel(ax, dateLabels{i}, rowLabelX);
                            set(ax,'TickDir','out','Box','off');

                            if i==1
                                title(ax, sprintf('Temperature Profile (dd/mm) (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);

                        outName = sprintf('TemperatureProfile_Dated_%s_%s.jpg', blockTag, blockRangeStr);
                        exportgraphics(figTPD, fullfile(outputFolderPath, outName),'Resolution',600);
                        close(figTPD);
                    end
                end

                % ----------------------------------------------------------
                % Light distribution and SD plot
                % ----------------------------------------------------------
                if chkLightDist.Value
                    blockLight = binnedLight(daysIdx,:);
                    [hourMean, hourSD] = hourlyMeanSD_fromBlock(blockLight, timeAxis);

                    figLD = figure('Visible','off','Color','w');
                    area(0:23, hourMean, 'FaceColor',[0 0.5 0.5]);
                    title(sprintf('Light Distribution (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                    xlabel('Hour of Day (0 = Midnight)','FontSize',14);
                    ylabel('Average Light (LUX)','FontSize',14);
                    xlim([0 23]); xticks(0:1:23);
                    set(gca,'TickDir','out','FontSize',11,'Box','off');

                    outName = sprintf('LightDistribution_Mean_%s_%s.jpg', blockTag, blockRangeStr);
                    exportgraphics(figLD, fullfile(outputFolderPath, outName),'Resolution',600);
                    close(figLD);

                    figLDSD = figure('Visible','off','Color','w');
                    ax = gca; hold(ax,'on');
                    xH = 0:23;

                    sdLower = max(0, hourMean - hourSD);
                    sdUpper = hourMean + hourSD;

                    bandX = [xH, fliplr(xH)];
                    bandY = [sdLower; flipud(sdUpper)]';
                    fill(ax, bandX, bandY, [0 0.5 0.5], 'FaceAlpha',0.2, 'EdgeColor','none');
                    plot(ax, xH, hourMean, 'LineWidth',2, 'Color',[0 0.5 0.5]);

                    title(sprintf('Light Distribution (Mean ± SD) (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                    xlabel('Hour of Day (0 = Midnight)','FontSize',14);
                    ylabel('Light (LUX)','FontSize',14);
                    xlim([0 23]); xticks(0:1:23);
                    set(ax,'TickDir','out','FontSize',11,'Box','off');
                    hold(ax,'off');

                    outName = sprintf('LightDistribution_MeanSD_%s_%s.jpg', blockTag, blockRangeStr);
                    exportgraphics(figLDSD, fullfile(outputFolderPath, outName),'Resolution',600);
                    close(figLDSD);
                end

                % ----------------------------------------------------------
                % Combined profile
                % ----------------------------------------------------------
                if chkCombined.Value
                    figCP = figure('Visible','off','Color','w');
                    set(figCP,'Position',[100 100 1200 600]);

                    for i = 1:nThisBlock
                        d = daysIdx(i);

                        axMain = subplot(nThisBlock,1,i);
                        hold(axMain,'on');
                        shiftAxisRight(axMain, leftMarginFrac);

                        axMain.XAxisLocation = 'origin';
                        axMain.XLim       = [0 1440];
                        axMain.XTick      = 0:240:1440;

                        if i ~= nThisBlock
                            axMain.XTickLabel = [];
                        else
                            axMain.XTickLabel = {'00:00','04:00','08:00','12:00','16:00','20:00','24:00'};
                        end

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

                        if anyDatesGlobal
                            addRowLabel(axMain, dateLabels{i}, rowLabelX);
                        else
                            addRowLabel(axMain, sprintf('Day %d', d), rowLabelX);
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
                        yticks(axTemp,[0 globalMinTemp, globalMaxTemp]);
                        ylabel(axTemp,'°C','FontSize',12);

                        linkaxes([axMain,axTemp],'x');

                        if i==1
                            title(axMain, sprintf('Combined Profile (%s)', strrep(blockRangeStr,'_',' ')),'FontSize',16);
                        end

                        hold(axMain,'off');
                        hold(axTemp,'off');
                    end

                    xlabel(axMain,'Time of Day','FontSize',12);

                    outName = sprintf('CombinedProfile_%s_%s.jpg', blockTag, blockRangeStr);
                    exportgraphics(figCP, fullfile(outputFolderPath, outName),'Resolution',600);
                    close(figCP);
                end
            end

            % --------------------------------------------------------------
            % Write Excel workbook
            % --------------------------------------------------------------
            setStatus('Writing Excel outputs...', 0.90);

            summaryFile = fullfile(outputFolderPath,'Participant_Results.xlsx');

            if chkExcelSummary.Value
                setStatus('Writing Excel: Summary sheet...', 0.91);
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
                setStatus('Writing Excel: Metrics sheet...', 0.93);

                daysPerFigureReported = daysPerFigure;
                if daysPerFigureReported == 0
                    daysPerFigureReported = NaN;
                end

                MetricsTable = table( ...
                    intervalMins, ...
                    lightThreshold, ...
                    logical(chkCompleteOnly.Value), ...
                    string(ddlAxisStyle.Value), ...
                    daysPerFigureReported, ...
                    ISvalue, ...
                    IVvalue, ...
                    string(usedSheet), ...
                    'VariableNames',{ ...
                        'SamplingInterval_Minutes', ...
                        'LightThreshold_Lux', ...
                        'CompleteDaysOnly', ...
                        'AxisStyle', ...
                        'DaysPerFigure_NaNmeansAll', ...
                        'InterdailyStability_Hourly', ...
                        'IntradailyVariability_Hourly', ...
                        'DataSheetUsed'} );
                writetable(MetricsTable, summaryFile, 'Sheet','Metrics');
            end

            setStatus('Writing Excel: Definitions sheet...', 0.94);
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

            %------------------------------------------------------------------
            % Compile PowerPoint (simplified + robust; includes subfolders)
            %------------------------------------------------------------------
            if chkPowerPoint.Value
                setStatus('Creating PowerPoint...', 0.96);

                if ~exist('mlreportgen.ppt.Presentation','class')
                    uialert(fig,'Report Generator toolbox not found. Cannot export PowerPoint.','PowerPoint Error');
                else
                    import mlreportgen.ppt.*;

                    pptFile = fullfile(outputFolderPath,'AllFigures_Report.pptx');
                    ppt     = Presentation(pptFile);
                    open(ppt);

                    jpgFiles = dir(fullfile(outputFolderPath,'**','*.jpg'));

                    if ~isempty(jpgFiles)
                        % Deterministic ordering: sort by full path as STRING (no lower() call)
                        fullPaths = string(fullfile({jpgFiles.folder}, {jpgFiles.name}));
                        [~, ord] = sort(fullPaths);
                        jpgFiles = jpgFiles(ord);

                        for k = 1:numel(jpgFiles)
                            slide = add(ppt,'Title and Content');

                            % Title: last folder name (if subfolder) + file base name
                            [~, baseName, ~] = fileparts(jpgFiles(k).name);

                            if strcmp(jpgFiles(k).folder, outputFolderPath)
                                titleStr = baseName;
                            else
                                [~, lastFolder] = fileparts(jpgFiles(k).folder);
                                titleStr = sprintf('%s - %s', lastFolder, baseName);
                            end

                            titleStr = strrep(titleStr, '_', ' ');
                            replace(slide,'Title', char(string(titleStr)));

                            picPath = fullfile(jpgFiles(k).folder, jpgFiles(k).name);
                            replace(slide,'Content', Picture(char(string(picPath))));
                        end
                    end

                    close(ppt);
                end
            end

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
        lblStatus.Text = char(string(msg));
        if ~isempty(prog) && isvalid(prog)
            prog.Message = char(string(msg));
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
    % Helper: add horizontal row label without squeezing axis
    % ---------------------------------------------------------------------
    function addRowLabel(ax, labelStr, xOffset)
        if nargin < 3 || isempty(xOffset)
            xOffset = -0.02;
        end
        text(ax, xOffset, 0.5, char(string(labelStr)), ...
            'Units','normalized', ...
            'HorizontalAlignment','right', ...
            'VerticalAlignment','middle', ...
            'Rotation',0, ...
            'FontSize',11);
    end

    % ---------------------------------------------------------------------
    % Helper: shift an axis to the right to create left margin buffer
    % ---------------------------------------------------------------------
    function shiftAxisRight(ax, leftMarginFrac)
        if nargin < 2 || isempty(leftMarginFrac)
            leftMarginFrac = 0.06;
        end
        pos = ax.Position;
        pos(1) = pos(1) + leftMarginFrac;
        pos(3) = max(0.01, pos(3) - leftMarginFrac);
        ax.Position = pos;
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
        % Prefer sheetnames, fall back to xlsfinfo for older installs
        try
            sheets = sheetnames(xlsxPath);
        catch
            [~, sheets] = xlsfinfo(xlsxPath);
            sheets = string(sheets);
        end

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
            m = nan(1, 1);
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
        [~, nBins] = size(dailyMat);

        timeAxisLocal = (0:nBins-1) * intervalMins;
        hourIdx = floor(timeAxisLocal/60) + 1;
        hourIdx(hourIdx < 1) = 1;
        hourIdx(hourIdx > 24) = 24;

        hourlyMat = nan(size(dailyMat,1), 24);
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

        meanByHour = mean(hourlyMat, 1, 'omitnan');
        p = 24;

        denom = sum((xValid - grandMean).^2, 'omitnan');
        numer = p * sum((meanByHour - grandMean).^2, 'omitnan');

        if denom <= 0 || isnan(denom)
            ISvalue = nan;
        else
            ISvalue = numer / denom;
        end

        xGrid = x;
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

    % ---------------------------------------------------------------------
    % Helper: hourly mean & SD for a block (days x bins)
    % ---------------------------------------------------------------------
    function [hourMean, hourSD] = hourlyMeanSD_fromBlock(lightBlock, timeAxisVals)
        binHour = floor(timeAxisVals/60);
        binHour(binHour > 23) = 23;

        hourMean = nan(24,1);
        hourSD   = nan(24,1);

        for h = 0:23
            vals = lightBlock(:, binHour==h);
            v = vals(:);
            hourMean(h+1) = mean(v,'omitnan');
            hourSD(h+1)   = std(v,'omitnan');
        end

        hourMean(isnan(hourMean)) = 0;
        hourSD(isnan(hourSD))     = 0;
    end

    % ---------------------------------------------------------------------
    % Helper: hourly mean & SD for a single day
    % ---------------------------------------------------------------------
    function [hourMean, hourSD] = hourlyMeanSD_fromSingleDay(lightDay, timeAxisVals)
        binHour = floor(timeAxisVals/60);
        binHour(binHour > 23) = 23;

        hourMean = nan(24,1);
        hourSD   = nan(24,1);

        for h = 0:23
            v = lightDay(binHour==h);
            hourMean(h+1) = mean(v,'omitnan');
            hourSD(h+1)   = std(v,'omitnan');
        end

        hourMean(isnan(hourMean)) = 0;
        hourSD(isnan(hourSD))     = 0;
    end

    % ---------------------------------------------------------------------
    % Tiny helper: ternary operator for strings
    % ---------------------------------------------------------------------
    function out = ternary(cond, a, b)
        if cond
            out = a;
        else
            out = b;
        end
        out = char(string(out));
    end

end