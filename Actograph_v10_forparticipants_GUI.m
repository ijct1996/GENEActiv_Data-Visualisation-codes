function Actograph_v10_forparticipants_GUI
% -------------------------------------------------------------------------
% Actograph Analysis and Summary GUI for GENEActiv Participants (v10.5)
% See README files for usage, input format, outputs, and interpretation notes.
%
% Last update: 04 February 2026 (IJT)
% Change summary: Double-plotted actograms and output feature optimisation.
%
% © 2025–2026 Isaiah J. Ting, Lall Lab
% -------------------------------------------------------------------------

    clc; close all;

    if exist('uifigure','file') ~= 2
        error(['Cannot find uifigure. This requires MATLAB with App Designer UI components. ', ...
               'If you are running a restricted environment, install/enable the required UI components.']);
    end

    % ---------------------------------------------------------------------
    % Universal figure defaults
    % ---------------------------------------------------------------------
    set(0, ...
        'DefaultFigureUnits','pixels', ...
        'DefaultFigurePosition',[100 100 1200 600], ...
        'DefaultFigurePaperUnits','inches', ...
        'DefaultFigurePaperPositionMode','auto');

    % ---------------------------------------------------------------------
    % Shared analysis state (PREDECLARE for nested function visibility)
    % ---------------------------------------------------------------------
    outputFolderPath   = "";
    useVectorPDF       = false;

    % Data/meta
    rawTable = table();
    meta     = struct();
    usedSheet = "";
    timestamps = datetime.empty(0,1);
    timestampsLabel = datetime.empty(0,1);
    activityCounts = [];
    lightLevels    = [];
    temperature    = [];

    % Sampling/grid
    epochSec     = NaN;
    intervalMins = NaN;
    binsPerDay   = NaN;
    totalDays    = 0;
    day0         = datetime.empty(0,1);

    % Day labelling
    dayStartDatesNum   = datetime.empty(0,1);
    dayStartDatesLabel = datetime.empty(0,1);
    weekDayNames = {};
    dateLabels   = {};
    dateISO      = {};

    % Binned daily matrices
    binnedActivity    = [];
    binnedLight       = [];
    binnedTemperature = [];

    % Metrics
    totalActivityPerDay  = [];
    hoursInLightPerDay   = [];
    minTempPerDay        = [];
    maxTempPerDay        = [];
    L5TimeStrings        = strings(0,1);
    M10TimeStrings       = strings(0,1);
    L5Mean               = [];
    M10Mean              = [];
    ISvalue              = NaN;
    IVvalue              = NaN;

    % Double-plot matrices (0–48 h)
    dpAct  = [];
    dpLux  = [];
    dpTemp = [];
    xHours48 = [];

    % Block naming
    blockStartStr = "";
    blockEndStr   = "";
    blockRangeISO = "";

    % Styling
    missingBG = [0.96 0.96 0.96];   % missing epochs background
    darkFill  = [0.35 0.35 0.35];   % darker grey for non-light
    lightFill = [0.9290 0.6940 0.1250];

    % Row label placement (no axes shifting)
    labelX = -0.012;
    labelFont = 11;

    % Thresholds/options
    lightThreshold = 5;
    tempHasAny = false;

    % Global temperature scaling (whole dataset)
    tempGlobalMinTick = NaN;
    tempGlobalMaxTick = NaN;
    tempGlobalTicks   = [];

    % Label modes struct for y-labelled outputs
    labelModes = struct('tag',{},'labels',{});

    % PPT export queue (deterministic slide ordering)
    pptQueuePaths  = strings(0,1);
    pptQueueTitles = strings(0,1);

    % Progress dialog handle
    prog = [];

    % ---------------------------------------------------------------------
    % Build GUI (scrollable when available)
    % ---------------------------------------------------------------------
    fig = uifigure('Name','Actograph Analysis','Position',[220 120 980 780],'Resize','on');

    rootPanel = uipanel(fig,'Units','normalized','Position',[0 0 1 1]);
    if isprop(rootPanel,'Scrollable')
        rootPanel.Scrollable = 'on';
    end

    mainGrid = uigridlayout(rootPanel,[11 3]);
    mainGrid.ColumnWidth   = {170,'1x',170};
    mainGrid.RowHeight     = {'fit','fit','fit','fit','fit','fit','fit','fit','fit','fit','1x'};
    mainGrid.Padding       = [12 12 12 12];
    mainGrid.RowSpacing    = 10;
    mainGrid.ColumnSpacing = 10;

    % Row 1: Input file
    lblIn = uilabel(mainGrid,'Text','Input file:','HorizontalAlignment','left');
    lblIn.Layout.Row = 1; lblIn.Layout.Column = 1;

    txtInputFile = uieditfield(mainGrid,'text');
    txtInputFile.Layout.Row = 1; txtInputFile.Layout.Column = 2;

    btnBrowseIn = uibutton(mainGrid,'push','Text','Browse...','ButtonPushedFcn',@selectInputFile);
    btnBrowseIn.Layout.Row = 1; btnBrowseIn.Layout.Column = 3;

    % Row 2: Output folder
    lblOut = uilabel(mainGrid,'Text','Output folder:','HorizontalAlignment','left');
    lblOut.Layout.Row = 2; lblOut.Layout.Column = 1;

    txtOutputFolder = uieditfield(mainGrid,'text');
    txtOutputFolder.Layout.Row = 2; txtOutputFolder.Layout.Column = 2;

    btnBrowseOut = uibutton(mainGrid,'push','Text','Browse...','ButtonPushedFcn',@selectOutputFolder);
    btnBrowseOut.Layout.Row = 2; btnBrowseOut.Layout.Column = 3;

    % Row 3: Timezone (labels only)
    lblTZ = uilabel(mainGrid,'Text','Timezone (labels only):','HorizontalAlignment','left');
    lblTZ.Layout.Row = 3; lblTZ.Layout.Column = 1;

    txtTimeZone = uieditfield(mainGrid,'text');
    txtTimeZone.Layout.Row = 3; txtTimeZone.Layout.Column = 2;
    txtTimeZone.Placeholder = 'e.g. Europe/London (leave blank for naive local time)';

    tzHint = uilabel(mainGrid,'Text','IANA: Continent/City','HorizontalAlignment','left');
    tzHint.Layout.Row = 3; tzHint.Layout.Column = 3;
    tzHint.FontColor = [0.35 0.35 0.35];

    % Row 4: Light threshold
    lblThr = uilabel(mainGrid,'Text','Light threshold (lux):','HorizontalAlignment','left');
    lblThr.Layout.Row = 4; lblThr.Layout.Column = 1;

    numThreshold = uieditfield(mainGrid,'numeric','Value',5,'Limits',[0 Inf], ...
        'RoundFractionalValues',true);
    numThreshold.Layout.Row = 4; numThreshold.Layout.Column = 2;

    thrHint = uilabel(mainGrid,'Text','Used for light shading + hours-in-light','HorizontalAlignment','left');
    thrHint.Layout.Row = 4; thrHint.Layout.Column = 3;
    thrHint.FontColor = [0.35 0.35 0.35];

    % Row 5: Axis style
    lblAxis = uilabel(mainGrid,'Text','Axis style:','HorizontalAlignment','left');
    lblAxis.Layout.Row = 5; lblAxis.Layout.Column = 1;

    ddlAxisStyle = uidropdown(mainGrid,'Items',{'Days only','Dated only','Both'},'Value','Both');
    ddlAxisStyle.Layout.Row = 5; ddlAxisStyle.Layout.Column = 2;

    axHint = uilabel(mainGrid,'Text','Controls y-axis labels for y-labelled outputs','HorizontalAlignment','left');
    axHint.Layout.Row = 5; axHint.Layout.Column = 3;
    axHint.FontColor = [0.35 0.35 0.35];

    % Row 6: Outputs panel
    pnlOutputs = uipanel(mainGrid,'Title','Select Outputs');
    pnlOutputs.Layout.Row = 6; pnlOutputs.Layout.Column = [1 3];

    outGrid = uigridlayout(pnlOutputs,[4 3]);
    outGrid.ColumnWidth = {'1x','1x','1x'};
    outGrid.RowHeight   = {'fit','fit','fit','fit'};
    outGrid.Padding = [10 10 10 10];
    outGrid.RowSpacing = 6;
    outGrid.ColumnSpacing = 16;

    chkAll = uicheckbox(outGrid,'Text','All outputs','Value',true,'ValueChangedFcn',@toggleAllOutputs);
    chkAll.Layout.Row = 1; chkAll.Layout.Column = 1;

    chkActivityProfDP = uicheckbox(outGrid,'Text','Activity profile (0–48 h)','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkActivityProfDP.Layout.Row = 2; chkActivityProfDP.Layout.Column = 1;

    chkHeatmapDP = uicheckbox(outGrid,'Text','Activity heatmap (0–48 h)','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkHeatmapDP.Layout.Row = 2; chkHeatmapDP.Layout.Column = 2;

    chkTempProfDP = uicheckbox(outGrid,'Text','Temperature profile (0–48 h)','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkTempProfDP.Layout.Row = 2; chkTempProfDP.Layout.Column = 3;

    chkCombinedDP = uicheckbox(outGrid,'Text','Combined profile (0–48 h)','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkCombinedDP.Layout.Row = 3; chkCombinedDP.Layout.Column = 1;

    chkDailyActivity = uicheckbox(outGrid,'Text','Daily activity totals','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkDailyActivity.Layout.Row = 3; chkDailyActivity.Layout.Column = 2;

    chkLowActivity = uicheckbox(outGrid,'Text','Low-activity call-outs','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkLowActivity.Layout.Row = 3; chkLowActivity.Layout.Column = 3;

    chkLightTracker = uicheckbox(outGrid,'Text','Daily light tracker (0–24 h)','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkLightTracker.Layout.Row = 4; chkLightTracker.Layout.Column = 1;

    chkLightDist = uicheckbox(outGrid,'Text','Light distribution (0–24 h)','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkLightDist.Layout.Row = 4; chkLightDist.Layout.Column = 2;

    chkExcel = uicheckbox(outGrid,'Text','Excel outputs','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkExcel.Layout.Row = 4; chkExcel.Layout.Column = 3;

    % Row 7: PowerPoint
    chkPowerPoint = uicheckbox(mainGrid,'Text','PowerPoint (compile outputs)','Value',true,'ValueChangedFcn',@syncAllCheckbox);
    chkPowerPoint.Layout.Row = 7; chkPowerPoint.Layout.Column = [1 3];

    % Row 8: Close GUI option
    chkCloseGUI = uicheckbox(mainGrid,'Text','Close GUI when finished','Value',true);
    chkCloseGUI.Layout.Row = 8; chkCloseGUI.Layout.Column = [1 3];

    % Row 9: Run
    btnRun = uibutton(mainGrid,'push','Text','Run','FontSize',14,'ButtonPushedFcn',@runAnalysis);
    btnRun.Layout.Row = 9; btnRun.Layout.Column = [1 3];

    % Row 10: Status
    lblStatus = uilabel(mainGrid,'Text','Ready','HorizontalAlignment','left');
    lblStatus.Layout.Row = 10; lblStatus.Layout.Column = [1 3];

    % Row 11: Error report panel
    errPanel = uipanel(mainGrid,'Title','Log / Error report');
    errPanel.Layout.Row = 11; errPanel.Layout.Column = [1 3];
    errGrid = uigridlayout(errPanel,[1 1]);
    errGrid.Padding = [6 6 6 6];

    txtError = uitextarea(errGrid,'Editable','off','Value',{'(No messages yet)'});
    txtError.FontName = 'Courier New';
    txtError.FontSize = 11;

    % ---------------------------------------------------------------------
    % Callbacks
    % ---------------------------------------------------------------------
    function selectInputFile(~,~)
        [file,path] = uigetfile('*.xlsx','Select Excel file');
        if isequal(file,0), return; end
        txtInputFile.Value = fullfile(path,file);
    end

    function selectOutputFolder(~,~)
        folder = uigetdir('','Select output folder');
        if isequal(folder,0), return; end
        txtOutputFolder.Value = folder;
    end

    function toggleAllOutputs(src,~)
        v = logical(src.Value);
        chkActivityProfDP.Value = v;
        chkHeatmapDP.Value      = v;
        chkTempProfDP.Value     = v;
        chkCombinedDP.Value     = v;
        chkDailyActivity.Value  = v;
        chkLowActivity.Value    = v;
        chkLightTracker.Value   = v;
        chkLightDist.Value      = v;
        chkExcel.Value          = v;
        chkPowerPoint.Value     = v;
    end

    function syncAllCheckbox(~,~)
        allSelected = all([ ...
            chkActivityProfDP.Value, chkHeatmapDP.Value, chkTempProfDP.Value, chkCombinedDP.Value, ...
            chkDailyActivity.Value, chkLowActivity.Value, chkLightTracker.Value, chkLightDist.Value, ...
            chkExcel.Value, chkPowerPoint.Value ]);
        chkAll.Value = logical(allSelected);
    end

    % ---------------------------------------------------------------------
    % Main analysis
    % ---------------------------------------------------------------------
    function runAnalysis(~,~)
        try
            clearLog();

            % Reset PPT queue every run
            pptQueuePaths  = strings(0,1);
            pptQueueTitles = strings(0,1);

            try
                evalin('base','clc;');
            catch
            end

            if ~any([chkActivityProfDP.Value, chkHeatmapDP.Value, chkTempProfDP.Value, chkCombinedDP.Value, ...
                     chkDailyActivity.Value, chkLowActivity.Value, chkLightTracker.Value, chkLightDist.Value, ...
                     chkExcel.Value, chkPowerPoint.Value])
                uialert(fig,'Please select at least one output.','Output Selection Error');
                return;
            end

            inputFilePath = string(txtInputFile.Value);
            if inputFilePath == "" || ~isfile(inputFilePath)
                uialert(fig,'Please select a valid input file.','Input Error');
                return;
            end

            outputFolderPath = string(txtOutputFolder.Value);
            if outputFolderPath == ""
                uialert(fig,'Please select an output folder.','Output Error');
                return;
            end
            if ~isfolder(outputFolderPath)
                mkdir(outputFolderPath);
            end

            prog = uiprogressdlg(fig,'Title','Running analysis','Message','Starting...','Value',0,'Cancelable',false);

            setStatus('Reading Excel workbook...', 0.05);
            [rawTable, meta, usedSheet] = loadRawData(inputFilePath);

            setStatus('Parsing timestamps...', 0.10);
            timestamps = parseTimestampsSoft(rawTable.(meta.timeVar));
            validTS = ~isnat(timestamps);
            if nnz(validTS) < 10
                error('Too few valid timestamps after parsing. Check the Time stamp column.');
            end

            activityCounts = double(rawTable.(meta.activityVar));
            lightLevels    = double(rawTable.(meta.lightVar));

            % Temperature optional
            if meta.hasTemp
                temperature = double(rawTable.(meta.tempVar));
            else
                temperature = nan(height(rawTable),1);
                appendLog('NOTE: Temperature column not found. Temperature outputs will be skipped; combined plot will omit temperature.');
            end

            timestamps     = timestamps(validTS);
            activityCounts = activityCounts(validTS);
            lightLevels    = lightLevels(validTS);
            temperature    = temperature(validTS);

            setStatus('Sorting by time...', 0.14);
            [timestamps, sIdx] = sort(timestamps);
            activityCounts = activityCounts(sIdx);
            lightLevels    = lightLevels(sIdx);
            temperature    = temperature(sIdx);

            % Timezone labels only
            tzIn = strtrim(string(txtTimeZone.Value));
            tzActive = false;
            timestampsLabel = timestamps;

            if tzIn ~= ""
                setStatus('Validating timezone (labels only)...', 0.16);
                try
                    timestampsLabel.TimeZone = char(tzIn);
                    tzActive = true;
                catch
                    error('Timezone "%s" not recognised. Use IANA format like "Europe/London".', tzIn);
                end
            end

            lightThreshold = numThreshold.Value;

            % Infer epoch
            setStatus('Inferring sampling epoch...', 0.20);
            dtSec = seconds(diff(timestamps));
            dtSec = dtSec(dtSec > 0 & isfinite(dtSec));
            if isempty(dtSec)
                error('Could not infer sampling interval from timestamps.');
            end
            epochSecRaw = round(median(dtSec));
            epochSec = snapEpochToDayDivisor(max(60, epochSecRaw));
            if mod(86400, epochSec) ~= 0
                epochSec = 60;
            end
            intervalMins = epochSec / 60;
            binsPerDay   = 86400 / epochSec;

            % Regularise to fixed epoch grid (NaN for missing)
            setStatus('Regularising to fixed epoch grid (NaN for missing)...', 0.28);

            tt = timetable(timestamps(:), activityCounts(:), lightLevels(:), temperature(:), ...
                'VariableNames',{'Activity','Light','Temp'});
            tt = sortrows(tt);

            ttR = retime(tt,'regular','mean','TimeStep',seconds(epochSec));

            day0 = dateshift(ttR.Time(1),'start','day');
            dayEnd = dateshift(ttR.Time(end),'start','day') + days(1) - seconds(epochSec);

            tGrid = (day0 : seconds(epochSec) : dayEnd)';
            ttG = retime(ttR, tGrid, 'mean');

            nTotal = numel(ttG.Time);
            totalDays = floor(nTotal / binsPerDay);
            if totalDays < 1
                error('No complete day grid could be formed from the recording.');
            end
            nUse = totalDays * binsPerDay;

            aVec = ttG.Activity(1:nUse);
            lVec = ttG.Light(1:nUse);
            tVec = ttG.Temp(1:nUse);

            binnedActivity    = reshape(aVec, binsPerDay, totalDays)'; % days x bins
            binnedLight       = reshape(lVec, binsPerDay, totalDays)';
            binnedTemperature = reshape(tVec, binsPerDay, totalDays)';

            dayStartDatesNum = day0 + days(0:(totalDays-1))';

            if tzActive
                dayStartDatesLabel = dayStartDatesNum;
                dayStartDatesLabel.TimeZone = char(tzIn);
            else
                dayStartDatesLabel = dayStartDatesNum;
            end

            weekDayNames = cellstr(day(dayStartDatesLabel,'name'));
            dateLabels   = cellstr(string(dayStartDatesLabel,'dd/MM'));
            dateISO      = cellstr(string(dayStartDatesLabel,'yyyy-MM-dd'));

            blockStartStr = string(dayStartDatesLabel(1),'yyyy-MM-dd');
            blockEndStr   = string(dayStartDatesLabel(end),'yyyy-MM-dd');
            blockRangeISO = sprintf('%s_to_%s', blockStartStr, blockEndStr);

            % Summary metrics
            setStatus('Computing daily summary metrics...', 0.36);
            totalActivityPerDay = sum(binnedActivity,2,'omitnan');

            validLux = ~isnan(binnedLight);
            hoursInLightPerDay = sum((binnedLight > lightThreshold) & validLux, 2) * (epochSec/3600);

            minTempPerDay = min(binnedTemperature,[],2,'omitnan');
            maxTempPerDay = max(binnedTemperature,[],2,'omitnan');

            % L5 / M10
            setStatus('Computing L5 and M10...', 0.40);
            L5Bins  = max(1, round(5*3600/epochSec));
            M10Bins = max(1, round(10*3600/epochSec));
            minFracValid = 0.90;

            L5StartNum  = NaT(totalDays,1);
            M10StartNum = NaT(totalDays,1);
            L5Mean = nan(totalDays,1);
            M10Mean = nan(totalDays,1);

            for d = 1:totalDays
                sig = binnedActivity(d,:);
                cL5  = slidingMeanNan(sig, L5Bins,  minFracValid);
                cM10 = slidingMeanNan(sig, M10Bins, minFracValid);

                if ~all(isnan(cL5))
                    [L5Mean(d), idx5] = min(cL5, [], 'omitnan');
                    L5StartNum(d) = dayStartDatesNum(d) + minutes((idx5-1)*intervalMins);
                end
                if ~all(isnan(cM10))
                    [M10Mean(d), idx10] = max(cM10, [], 'omitnan');
                    M10StartNum(d) = dayStartDatesNum(d) + minutes((idx10-1)*intervalMins);
                end
            end

            L5TimeStrings  = strings(totalDays,1);
            M10TimeStrings = strings(totalDays,1);
            for d = 1:totalDays
                if ~isnat(L5StartNum(d)),  L5TimeStrings(d)  = string(L5StartNum(d),'HH:mm:ss'); end
                if ~isnat(M10StartNum(d)), M10TimeStrings(d) = string(M10StartNum(d),'HH:mm:ss'); end
            end

            % IS/IV
            setStatus('Computing IS and IV (hourly)...', 0.46);
            [ISvalue, IVvalue] = computeISIV_hourlyFromDaily(binnedActivity, intervalMins);

            % Double-plot matrices
            setStatus('Preparing 0–48 h matrices...', 0.50);
            dpBins = 2*binsPerDay;

            dpAct  = nan(totalDays, dpBins);
            dpLux  = nan(totalDays, dpBins);
            dpTemp = nan(totalDays, dpBins);

            for d = 1:totalDays
                dpAct(d,1:binsPerDay)  = binnedActivity(d,:);
                dpLux(d,1:binsPerDay)  = binnedLight(d,:);
                dpTemp(d,1:binsPerDay) = binnedTemperature(d,:);
                if d < totalDays
                    dpAct(d,binsPerDay+1:end)  = binnedActivity(d+1,:);
                    dpLux(d,binsPerDay+1:end)  = binnedLight(d+1,:);
                    dpTemp(d,binsPerDay+1:end) = binnedTemperature(d+1,:);
                end
            end

            xHours48 = (0:(dpBins-1)) * (epochSec/3600);

            % Export mode switch
            maxDPRowsForVector = 30;
            useVectorPDF = (totalDays > maxDPRowsForVector);

            % Row label font scaling
            labelFont = 11;
            if totalDays > 60, labelFont = 9; end
            if totalDays > 90, labelFont = 8; end

            % Label modes
            yLabsDay  = arrayfun(@(x)sprintf('Day %d',x), (1:totalDays)', 'UniformOutput',false);
            yLabsDate = dateLabels(:);
            labelModes = buildLabelModes(ddlAxisStyle.Value, yLabsDay, yLabsDate);

            % Temperature availability and global scaling
            tempHasAny = any(isfinite(dpTemp(:)) & ~isnan(dpTemp(:)));
            if tempHasAny
                vT = dpTemp(isfinite(dpTemp));
                tMinRaw = min(vT);
                tMaxRaw = max(vT);
                [tempGlobalMinTick, tempGlobalMaxTick] = outwardRoundTempTicks(tMinRaw, tMaxRaw, 0.1);
                tempGlobalTicks = [tempGlobalMinTick tempGlobalMaxTick];
            else
                tempGlobalMinTick = NaN;
                tempGlobalMaxTick = NaN;
                tempGlobalTicks = [];
            end

            % -----------------------------------------------------------------
            % Outputs
            % -----------------------------------------------------------------
            setStatus('Generating outputs...', 0.55);

            % Daily light tracker (0–24 h)
            if chkLightTracker.Value
                setStatus('Daily light tracker (0–24 h)...', 0.58);
                lightTrackerFolder = fullfile(outputFolderPath,'DailyLightTracker');
                if ~isfolder(lightTrackerFolder), mkdir(lightTrackerFolder); end

                timeAxis24 = (0:binsPerDay-1) * intervalMins;

                for d = 1:totalDays
                    dateLabelISO = string(dayStartDatesLabel(d),'yyyy-MM-dd');
                    dateLabelNice = string(dayStartDatesLabel(d),'dd MMM yyyy');

                    figLT = figure('Visible','off','Color','w','Position',[100 100 900 800]);

                    ax1 = subplot(3,1,1);
                    plot(ax1, timeAxis24, binnedLight(d,:), 'Color',[0 0.5 0.5],'LineWidth',1);
                    hold(ax1,'on');
                    yline(ax1, lightThreshold,':','Color',[0.5 0.5 0.5],'LineWidth',0.5);
                    hold(ax1,'off');
                    title(ax1, sprintf('Daily Light (Full Range) | %s', dateLabelNice), ...
                        'FontSize',14,'HorizontalAlignment','center');
                    xlim(ax1,[0 1440]);
                    ax1.XTick = [];
                    set(ax1,'TickDir','out','Box','off','FontSize',11);
                    ylabel(ax1,'Light (lux)','FontSize',12);

                    axM = subplot(3,1,2); hold(axM,'on');
                    [hourMean, hourSD] = hourlyMeanSD_fromSingleDay(binnedLight(d,:), timeAxis24);
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
                    xlabel(axM,'Hour of day','FontSize',12);
                    ylabel(axM,'Light (lux)','FontSize',12);
                    set(axM,'TickDir','out','Box','off','FontSize',11);
                    hold(axM,'off');

                    ax2 = subplot(3,1,3);
                    plot(ax2, timeAxis24, binnedLight(d,:), 'Color',[0 0.5 0.5],'LineWidth',1);
                    hold(ax2,'on');
                    yline(ax2, lightThreshold,':','Color',[0.5 0.5 0.5],'LineWidth',0.5);
                    hold(ax2,'off');
                    title(ax2,'Daily Light (0–10 lux)','FontSize',14,'HorizontalAlignment','center');
                    xlim(ax2,[0 1440]);
                    ylim(ax2,[0 10]);
                    set(ax2,'TickDir','out','Box','off','FontSize',11);
                    xlabel(ax2,'Time of day','FontSize',12);
                    ylabel(ax2,'Light (lux)','FontSize',12);
                    xticks(ax2,0:240:1440);
                    xticklabels(ax2,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});

                    outPath = fullfile(lightTrackerFolder, sprintf('DailyLightTracker_%s.jpg', dateLabelISO));
                    exportgraphics(figLT, outPath, 'Resolution',600);

                    if chkPowerPoint.Value
                        queuePPTImage(outPath, sprintf('Daily Light Tracker %s', dateLabelISO));
                    end

                    close(figLT);
                end
            end

            % Light distribution (0–24 h)
            if chkLightDist.Value
                setStatus('Light distribution (0–24 h)...', 0.62);
                timeAxis24 = (0:binsPerDay-1) * intervalMins;
                [hourMean, hourSD] = hourlyMeanSD_fromBlock(binnedLight, timeAxis24);

                figLD = figure('Visible','off','Color','w');
                area(0:23, hourMean, 'FaceColor',[0 0.5 0.5]);
                title(sprintf('Light Distribution (0–24 h) | %s to %s', blockStartStr, blockEndStr),'FontSize',16);
                xlabel('Hour of day','FontSize',14);
                ylabel('Mean light (lux)','FontSize',14);
                xlim([0 23]); xticks(0:1:23);
                set(gca,'TickDir','out','FontSize',11,'Box','off');

                outPath = fullfile(outputFolderPath, sprintf('LightDistribution_Mean_0-24h_%s.jpg', blockRangeISO));
                exportgraphics(figLD, outPath, 'Resolution',600);

                if chkPowerPoint.Value
                    queuePPTImage(outPath, 'Light Distribution Mean');
                end
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
                title(sprintf('Light Distribution (Mean ± SD) (0–24 h) | %s to %s', blockStartStr, blockEndStr),'FontSize',16);
                xlabel('Hour of day','FontSize',14);
                ylabel('Light (lux)','FontSize',14);
                xlim([0 23]); xticks(0:1:23);
                set(ax,'TickDir','out','FontSize',11,'Box','off');
                hold(ax,'off');

                outPath = fullfile(outputFolderPath, sprintf('LightDistribution_MeanSD_0-24h_%s.jpg', blockRangeISO));
                exportgraphics(figLDSD, outPath, 'Resolution',600);

                if chkPowerPoint.Value
                    queuePPTImage(outPath, 'Light Distribution Mean ± SD');
                end
                close(figLDSD);
            end

            % Activity profile (0–48 h)
            if chkActivityProfDP.Value
                setStatus('Activity profile (0–48 h)...', 0.66);
                for m = 1:numel(labelModes)
                    makeAndExportActivityProfile(labelModes(m).labels, labelModes(m).tag);
                end
            end

            % Temperature profile (0–48 h)
            if chkTempProfDP.Value
                if tempHasAny
                    setStatus('Temperature profile (0–48 h)...', 0.70);
                    for m = 1:numel(labelModes)
                        makeAndExportTemperatureProfile(labelModes(m).labels, labelModes(m).tag);
                    end
                else
                    appendLog('NOTE: Temperature profile requested but temperature data are missing or all NaN. Skipping temperature profile.');
                end
            end

            % Combined profile (0–48 h)
            if chkCombinedDP.Value
                if tempHasAny
                    setStatus('Combined profile (0–48 h)...', 0.74);
                    for m = 1:numel(labelModes)
                        makeAndExportCombinedProfile(labelModes(m).labels, labelModes(m).tag, true);
                    end
                else
                    appendLog('NOTE: Combined profile requested but temperature data are missing or all NaN. Export will omit temperature.');
                    for m = 1:numel(labelModes)
                        makeAndExportCombinedProfile(labelModes(m).labels, labelModes(m).tag, false);
                    end
                end
            end

            % Activity heatmap (0–48 h)
            if chkHeatmapDP.Value
                setStatus('Activity heatmap (0–48 h)...', 0.78);
                for m = 1:numel(labelModes)
                    makeAndExportHeatmap(labelModes(m).labels, labelModes(m).tag);
                end
            end

            % Daily totals
            if chkDailyActivity.Value
                setStatus('Daily activity totals...', 0.82);
                figDA = figure('Visible','off','Color','w');
                bar(1:totalDays, totalActivityPerDay, 'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');

                xticks(1:totalDays);
                if strcmp(ddlAxisStyle.Value,'Dated only')
                    xticklabels(dateLabels);
                    xlabel('Date','FontSize',14);
                else
                    xticklabels(arrayfun(@(x)sprintf('Day %d',x),1:totalDays,'UniformOutput',false));
                    xlabel('Day','FontSize',14);
                end

                title(sprintf('Total Activity by Day | %s to %s', blockStartStr, blockEndStr),'FontSize',16);
                ylabel('Total activity','FontSize',14);
                set(gca,'TickDir','out','FontSize',11,'Box','off');

                outPath = fullfile(outputFolderPath, sprintf('DailyActivity_%s.jpg', blockRangeISO));
                exportgraphics(figDA, outPath, 'Resolution',600);

                if chkPowerPoint.Value
                    queuePPTImage(outPath, 'Daily Activity Totals');
                end
                close(figDA);
            end

            % -----------------------------------------------------------------
            % Low-activity call-outs (UPDATED: complete-day threshold + complete-day-only plot)
            % -----------------------------------------------------------------
            if chkLowActivity.Value
                setStatus('Low-activity call-outs...', 0.85);

                % Define "complete day" based on wear-time/coverage (activity epochs present).
                % This prevents edge days (partial recordings) being flagged as low solely due to missing data.
                completeDayMinFrac = 0.95; % near-complete coverage
                validCounts = sum(~isnan(binnedActivity), 2);
                minValidBins = ceil(completeDayMinFrac * binsPerDay);
                isCompleteDay = (validCounts >= minValidBins);

                nComplete = nnz(isCompleteDay);
                if nComplete < 3
                    appendLog(sprintf('NOTE: Only %d complete day(s) at >= %.2f coverage. Low threshold will use all days.', ...
                        nComplete, completeDayMinFrac));
                    isCompleteDayForThresh = true(totalDays,1);
                else
                    isCompleteDayForThresh = isCompleteDay;
                end

                mu = mean(totalActivityPerDay(isCompleteDayForThresh), 'omitnan');
                sd = std(totalActivityPerDay(isCompleteDayForThresh), 'omitnan');
                lowThresh = mu - sd;

                lowIdxs = find(isCompleteDay & (totalActivityPerDay < lowThresh));

                % --- Plot 1: all days, but partial days grey and never labelled "Low"
                figLC = figure('Visible','off','Color','w');

                b = bar(1:totalDays, totalActivityPerDay, 'EdgeColor','none');
                b.FaceColor = 'flat';

                blue = [0 0.4470 0.7410];
                grey = [0.80 0.80 0.80];
                b.CData = repmat(blue, totalDays, 1);
                b.CData(~isCompleteDay,:) = repmat(grey, nnz(~isCompleteDay), 1);

                hold on;
                ymax = max(totalActivityPerDay,[],'omitnan');
                if isempty(ymax) || isnan(ymax) || ymax == 0, ymax = 1; end
                for idx = lowIdxs'
                    text(idx, totalActivityPerDay(idx) + 0.05*ymax, 'Low', 'Color','r', ...
                        'FontSize',12,'HorizontalAlignment','center');
                end
                hold off;

                xticks(1:totalDays);
                if strcmp(ddlAxisStyle.Value,'Dated only')
                    xticklabels(dateLabels);
                    xlabel('Date','FontSize',14);
                else
                    xticklabels(arrayfun(@(x)sprintf('Day %d',x),1:totalDays,'UniformOutput',false));
                    xlabel('Day','FontSize',14);
                end

                title(sprintf('Low-Activity Call-outs (threshold from complete days) | %s to %s', ...
                    blockStartStr, blockEndStr),'FontSize',16);
                ylabel('Total activity','FontSize',14);
                set(gca,'TickDir','out','FontSize',11,'Box','off');

                outPath = fullfile(outputFolderPath, sprintf('LowActivity_%s.jpg', blockRangeISO));
                exportgraphics(figLC, outPath, 'Resolution',600);

                if chkPowerPoint.Value
                    queuePPTImage(outPath, 'Low Activity Call-outs (Complete-day threshold)');
                end
                close(figLC);

                % --- Plot 2: complete days only (the "third output" you described)
                if nnz(isCompleteDay) >= 1
                    figLCC = figure('Visible','off','Color','w');

                    dayIdx = find(isCompleteDay);
                    bar(dayIdx, totalActivityPerDay(isCompleteDay), 'FaceColor', blue, 'EdgeColor','none');
                    hold on;

                    ymax2 = max(totalActivityPerDay(isCompleteDay),[],'omitnan');
                    if isempty(ymax2) || isnan(ymax2) || ymax2 == 0, ymax2 = 1; end

                    lowIdxsC = dayIdx(totalActivityPerDay(isCompleteDay) < lowThresh);
                    for k = 1:numel(lowIdxsC)
                        dIdx = lowIdxsC(k);
                        text(dIdx, totalActivityPerDay(dIdx) + 0.05*ymax2, 'Low', 'Color','r', ...
                            'FontSize',12,'HorizontalAlignment','center');
                    end
                    hold off;

                    xticks(dayIdx);
                    if strcmp(ddlAxisStyle.Value,'Dated only')
                        xticklabels(dateLabels(isCompleteDay));
                        xlabel('Date (complete days only)','FontSize',14);
                    else
                        xticklabels(arrayfun(@(x)sprintf('Day %d',x), dayIdx, 'UniformOutput',false));
                        xlabel('Day (complete days only)','FontSize',14);
                    end

                    title(sprintf('Low-Activity Call-outs (Complete Days Only) | %s to %s', ...
                        blockStartStr, blockEndStr),'FontSize',16);
                    ylabel('Total activity','FontSize',14);
                    set(gca,'TickDir','out','FontSize',11,'Box','off');

                    outPath2 = fullfile(outputFolderPath, sprintf('LowActivity_CompleteDaysOnly_%s.jpg', blockRangeISO));
                    exportgraphics(figLCC, outPath2, 'Resolution',600);

                    if chkPowerPoint.Value
                        queuePPTImage(outPath2, 'Low Activity Call-outs (Complete Days Only)');
                    end
                    close(figLCC);
                else
                    appendLog('NOTE: No complete days met the coverage threshold. Skipping the "complete days only" low-activity plot.');
                end
            end

            % Excel outputs
            if chkExcel.Value
                setStatus('Writing Excel outputs...', 0.90);
                summaryFile = fullfile(outputFolderPath,'Participant_Results.xlsx');

                SummaryTable = table( ...
                    dayStartDatesLabel, ...
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

                exportMode = "JPG_600dpi";
                if useVectorPDF
                    exportMode = "PDF_vector_for_large_0-48h_outputs";
                end

                tzReported = "(none)";
                if tzActive, tzReported = string(tzIn); end

                MetricsTable = table( ...
                    intervalMins, ...
                    epochSec, ...
                    lightThreshold, ...
                    string(ddlAxisStyle.Value), ...
                    string(tzReported), ...
                    string(exportMode), ...
                    ISvalue, ...
                    IVvalue, ...
                    string(usedSheet), ...
                    totalDays, ...
                    string(tempHasAny), ...
                    'VariableNames',{ ...
                        'SamplingInterval_Minutes', ...
                        'EpochSeconds_RegularGrid', ...
                        'LightThreshold_Lux', ...
                        'AxisStyle', ...
                        'TimeZone_LabelsOnly', ...
                        'ExportMode', ...
                        'InterdailyStability_Hourly', ...
                        'IntradailyVariability_Hourly', ...
                        'DataSheetUsed', ...
                        'TotalDaysAnalysed', ...
                        'TemperatureAvailable'} );
                writetable(MetricsTable, summaryFile, 'Sheet','Metrics');

                defs = {
                    'Date', 'Calendar date of recording day (label only)', 'Used to align outputs to calendar days (timezone affects labels only)';
                    'Weekday', 'Day of the week', 'Useful for weekday vs weekend comparisons';
                    'Day', 'Sequential index from start of analysed block', 'Day 1 is the first complete analysed day';
                    'TotalActivity', 'Sum of activity across the day on the regularised grid', 'Overall daily movement (missing epochs excluded)';
                    'HoursInLight', sprintf('Hours with Light > %.3g lux (within valid lux epochs)', lightThreshold), 'Daily light exposure above threshold';
                    'L5_StartTime', 'Clock time when the lowest mean 5-hour window begins', 'Rest-activity trough onset proxy';
                    'L5_Mean', 'Mean activity within the L5 window', 'Depth of the daily trough';
                    'M10_StartTime', 'Clock time when the highest mean 10-hour window begins', 'Peak activity onset proxy';
                    'M10_Mean', 'Mean activity within the M10 window', 'Strength of daily peak activity';
                    'MinTemperature', 'Minimum recorded temperature per day (°C)', 'Daily minimum (if temperature available)';
                    'MaxTemperature', 'Maximum recorded temperature per day (°C)', 'Daily maximum (if temperature available)';
                    'SamplingInterval_Minutes', 'Median inferred sampling interval from raw timestamps', 'Used to choose a robust epoch';
                    'EpochSeconds_RegularGrid', 'Epoch used after regularisation onto a fixed grid', 'Missing epochs remain NaN (plotted as blank)';
                    'MissingDataHandling', 'No interpolation; missing epochs remain NaN', 'Blank regions indicate missing data';
                    'CompleteDayTruncation', 'Analysis uses complete days on the fixed grid', 'Partial start/end days are not analysed';
                    'TimeZone_LabelsOnly', 'Timezone is applied to labels only', 'Numeric binning and metrics ignore timezone/DST';
                    '0-48hOutputs', 'Row i shows Day i (0–24 h) then Day i+1 (24–48 h)', 'Final row second half is blank';
                    'HeatmapScaling', 'Colour scale uses 5th–95th percentile of non-missing values', 'Improves visibility; not a fixed absolute scale';
                    'ExportMode', 'JPG 600 dpi or vector PDF for large 0–48 h figures', 'Vector export avoids very large raster files';
                    'InterdailyStability_Hourly', 'IS computed from hourly means across days', 'Higher IS indicates stronger day-to-day regularity';
                    'IntradailyVariability_Hourly', 'IV computed from successive hourly differences', 'Higher IV indicates greater fragmentation'
                };

                writecell([{'Term','Definition','Interpretation'}; defs], summaryFile, 'Sheet','Definitions');
            end

            % PowerPoint compilation
            if chkPowerPoint.Value
                setStatus('Creating PowerPoint...', 0.96);

                if ~exist('mlreportgen.ppt.Presentation','class')
                    uialert(fig,'Report Generator toolbox not found. Cannot export PowerPoint.','PowerPoint Error');
                else
                    origDir = pwd;
                    cd(outputFolderPath);
                    cObj = onCleanup(@() cd(origDir)); 

                    pptFileName = 'AllFigures_Report.pptx';
                    if isfile(pptFileName)
                        try delete(pptFileName); catch, end
                    end

                    ppt = mlreportgen.ppt.Presentation(pptFileName);
                    open(ppt);

                    nQueued = numel(pptQueuePaths);
                    nAdded = 0;

                    if nQueued == 0
                        appendLog('NOTE: PPT queue is empty. Falling back to recursive search for JPGs.');
                        jpgFiles = dir(fullfile(outputFolderPath,'**','*.jpg'));
                        for k = 1:numel(jpgFiles)
                            p = fullfile(jpgFiles(k).folder, jpgFiles(k).name);
                            addSlideWithImage(ppt, p, "");
                            nAdded = nAdded + 1;
                        end
                    else
                        for k = 1:nQueued
                            p = char(pptQueuePaths(k));
                            if ~isfile(p), continue; end
                            [~,~,ext] = fileparts(p);
                            if ~strcmpi(ext,'.jpg') && ~strcmpi(ext,'.jpeg') && ~strcmpi(ext,'.png')
                                continue;
                            end
                            t = "";
                            if k <= numel(pptQueueTitles), t = char(pptQueueTitles(k)); end
                            addSlideWithImage(ppt, p, t);
                            nAdded = nAdded + 1;
                        end
                    end

                    close(ppt);

                    tmpDir = fullfile(outputFolderPath,'ppt_tmp');
                    if isfolder(tmpDir)
                        try rmdir(tmpDir,'s'); catch, end
                    end

                    appendLog(sprintf('PPT created with %d slide(s) from exported figures.', nAdded));
                end
            end

            setStatus('Done.', 1.0);
            closeProgressDialog();

            if chkCloseGUI.Value && isvalid(fig)
                drawnow;
                pause(0.15);
                delete(fig);
            end

        catch ME
            closeProgressDialog();
            reportError(ME);
            uialert(fig, ME.message, 'Error');
            lblStatus.Text = 'Error encountered.';
        end
    end

    % ---------------------------------------------------------------------
    % DP Figure builders
    % ---------------------------------------------------------------------
    function makeAndExportActivityProfile(yLabels, labelTag)
        figAP = figure('Visible','off','Color','w');
        set(figAP,'Position',[80 60 1700 max(900, min(26000, 60*totalDays + 260))]);

        tLay = tiledlayout(totalDays,1,'TileSpacing','compact','Padding','loose');

        for i = 1:totalDays
            ax = nexttile(tLay,i);
            ax.Color = missingBG;
            hold(ax,'on');

            lux = dpLux(i,:);
            mValid = ~isnan(lux);
            mL = (lux > lightThreshold) & mValid;
            mD = (~mL) & mValid;

            y = dpAct(i,:);
            yMax = max(y,[],'omitnan');
            if isnan(yMax) || yMax == 0, yMax = 1; end
            yMax = yMax * 1.2;

            fillSegments(ax,mD,xHours48,0,yMax,darkFill,0.36);
            fillSegments(ax,mL,xHours48,0,yMax,lightFill,0.22);

            bar(ax,xHours48,y,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');

            ylim(ax,[0 yMax]);
            xlim(ax,[0 48]);

            if i ~= totalDays
                ax.XTickLabel = [];
            else
                ax.XTick = 0:6:48;
                xlabel(ax,'Time (h)','FontSize',12);
            end

            ax.YTick = [];
            set(ax,'TickDir','out','Box','off','FontSize',10);

            addRowLabelTiled(ax, yLabels{i}, labelX, labelFont);

            hold(ax,'off');
        end

        title(tLay, sprintf('Activity Profile (0–48 h) | %s to %s', blockStartStr, blockEndStr), ...
            'FontSize',16,'FontWeight','bold');

        baseOut = fullfile(outputFolderPath, sprintf('Actogram_Activity_0-48h_%s_%s', labelTag, blockRangeISO));
        outPath = exportFigureSmart(figAP, baseOut, useVectorPDF, false);
        if chkPowerPoint.Value
            queuePPTImage(outPath, sprintf('Activity Profile %s', labelTag));
        end
        close(figAP);
    end

    function makeAndExportTemperatureProfile(yLabels, labelTag)
        figTP = figure('Visible','off','Color','w');
        set(figTP,'Position',[80 60 1850 max(900, min(26000, 60*totalDays + 260))]);

        if ~tempHasAny || isempty(tempGlobalTicks)
            appendLog('NOTE: Temperature profile skipped because temperature is missing or all NaN.');
            close(figTP);
            return;
        end

        tMin = tempGlobalMinTick;
        tMax = tempGlobalMaxTick;
        tTicks = tempGlobalTicks;

        tLay = tiledlayout(totalDays,1,'TileSpacing','compact','Padding','loose');

        for i = 1:totalDays
            ax = nexttile(tLay,i);
            ax.Color = missingBG;
            hold(ax,'on');

            lux = dpLux(i,:);
            mValid = ~isnan(lux);
            mL = (lux > lightThreshold) & mValid;
            mD = (~mL) & mValid;

            fillSegments(ax,mD,xHours48,tMin,tMax,darkFill,0.36);
            fillSegments(ax,mL,xHours48,tMin,tMax,lightFill,0.22);

            bar(ax,xHours48,dpTemp(i,:),'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');

            ylim(ax,[tMin tMax]);
            xlim(ax,[0 48]);

            if i ~= totalDays
                ax.XTickLabel = [];
            else
                ax.XTick = 0:6:48;
                xlabel(ax,'Time (h)','FontSize',12);
            end

            ax.YAxisLocation = 'right';
            ax.YTick = tTicks;
            ax.YTickLabel = formatTickLabels(tTicks, '%.1f');
            set(ax,'TickDir','out','Box','off','FontSize',10);

            addRowLabelTiled(ax, yLabels{i}, labelX, labelFont);

            hold(ax,'off');
        end

        title(tLay, sprintf('Temperature Profile (0–48 h) | %s to %s', blockStartStr, blockEndStr), ...
            'FontSize',16,'FontWeight','bold');

        addRightSideLabel(figTP, 'Temperature (°C)');

        baseOut = fullfile(outputFolderPath, sprintf('Actogram_Temperature_0-48h_%s_%s', labelTag, blockRangeISO));
        outPath = exportFigureSmart(figTP, baseOut, useVectorPDF, false);
        if chkPowerPoint.Value
            queuePPTImage(outPath, sprintf('Temperature Profile %s', labelTag));
        end
        close(figTP);
    end

    function makeAndExportCombinedProfile(yLabels, labelTag, hasTempIn)
        figCP = figure('Visible','off','Color','w');
        set(figCP,'Position',[80 60 1850 max(900, min(26000, 60*totalDays + 260))]);

        hasTemp = logical(hasTempIn) && tempHasAny && ~isempty(tempGlobalTicks);

        if hasTemp
            tMin = tempGlobalMinTick;
            tMax = tempGlobalMaxTick;
            tTicks = tempGlobalTicks;
        end

        tLay = tiledlayout(totalDays,1,'TileSpacing','compact','Padding','loose');

        for i = 1:totalDays
            axMain = nexttile(tLay,i);
            axMain.Color = missingBG;
            hold(axMain,'on');

            lux = dpLux(i,:);
            mValid = ~isnan(lux);
            mL = (lux > lightThreshold) & mValid;
            mD = (~mL) & mValid;

            yA = dpAct(i,:);
            yMax = max(yA,[],'omitnan');
            if isnan(yMax) || yMax == 0, yMax = 1; end
            yMax = yMax * 1.2;

            fillSegments(axMain,mD,xHours48,0,yMax,darkFill,0.36);
            fillSegments(axMain,mL,xHours48,0,yMax,lightFill,0.22);

            if hasTemp
                yyaxis(axMain,'left');
                bar(axMain,xHours48,yA,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                ylim(axMain,[0 yMax]);
                axMain.YTick = [];
                axMain.YColor = [0 0 0];

                yyaxis(axMain,'right');
                plot(axMain,xHours48,dpTemp(i,:),'LineWidth',1.2,'Color',[0.8500 0.3250 0.0980]);
                ylim(axMain,[tMin tMax]);
                axMain.YTick = tTicks;
                axMain.YTickLabel = formatTickLabels(tTicks, '%.1f');
                axMain.YColor = [0.25 0.25 0.25];
            else
                bar(axMain,xHours48,yA,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                ylim(axMain,[0 yMax]);
                axMain.YTick = [];
            end

            xlim(axMain,[0 48]);

            if i ~= totalDays
                axMain.XTickLabel = [];
            else
                axMain.XTick = 0:6:48;
                xlabel(axMain,'Time (h)','FontSize',12);
            end

            set(axMain,'TickDir','out','Box','off','FontSize',10);
            addRowLabelTiled(axMain, yLabels{i}, labelX, labelFont);

            hold(axMain,'off');
        end

        if hasTemp
            ttl = sprintf('Combined Profile (Activity, Light, Temperature) (0–48 h) | %s to %s', blockStartStr, blockEndStr);
            addRightSideLabel(figCP, 'Temperature (°C)');
        else
            ttl = sprintf('Combined Profile (Activity, Light) (0–48 h) | %s to %s', blockStartStr, blockEndStr);
        end
        title(tLay, ttl, 'FontSize',16,'FontWeight','bold');

        baseOut = fullfile(outputFolderPath, sprintf('Actogram_Combined_0-48h_%s_%s', labelTag, blockRangeISO));
        exportFigureSmart(figCP, baseOut, true, true); % force PDF
        if chkPowerPoint.Value
            appendLog(sprintf('NOTE: Combined profile (%s) exported as PDF and skipped for PPT.', labelTag));
        end
        close(figCP);
    end

    function makeAndExportHeatmap(yLabels, labelTag)
        figHM = figure('Visible','off','Color','w');
        figHM.Position = [80 60 1700 max(800, min(22000, 24*totalDays + 320))];

        ax = axes(figHM);
        ax.Color = missingBG;

        hImg = imagesc(ax, xHours48, 1:totalDays, dpAct);
        axis(ax,'tight');
        ax.YDir = 'reverse';

        set(hImg,'AlphaData', ~isnan(dpAct));

        v = dpAct(~isnan(dpAct));
        if isempty(v)
            setClimSafe(ax,[0 1]);
        else
            cMin = prctile(v,5);
            cMax = prctile(v,95);
            if cMin == cMax, cMax = cMin + 1; end
            setClimSafe(ax,[cMin cMax]);
        end

        colormap(ax, blueAmberRedMap(256));

        cb = colorbar(ax,'eastoutside');
        cb.Label.String = 'Activity Intensity';
        cb.Label.FontSize = 12;

        yticks(ax,1:totalDays);
        yticklabels(ax,yLabels);

        xlim(ax,[0 48]);
        xt = 0:6:48;
        xticks(ax,xt);
        xticklabels(ax, arrayfun(@(z) sprintf('%g',z), xt, 'UniformOutput', false));

        xlabel(ax,'Time (h)','FontSize',12);
        title(ax, sprintf('Activity Heatmap (0–48 h) | %s to %s', blockStartStr, blockEndStr), ...
            'FontSize',16,'FontWeight','bold');

        ax.Box = 'off';
        ax.TickDir = 'out';
        hold(ax,'on');
        rectangle(ax, 'Position',[0 0.5 48 totalDays], 'EdgeColor',[0 0 0], 'LineWidth',0.8);
        hold(ax,'off');

        baseOut = fullfile(outputFolderPath, sprintf('Heatmap_Activity_0-48h_%s_%s', labelTag, blockRangeISO));
        outPath = exportFigureSmart(figHM, baseOut, useVectorPDF, false);
        if chkPowerPoint.Value
            queuePPTImage(outPath, sprintf('Activity Heatmap %s', labelTag));
        end
        close(figHM);
    end

    % ---------------------------------------------------------------------
    % Status + log helpers
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

    function clearLog()
        txtError.Value = {'(No messages yet)'};
    end

    function appendLog(msg)
        if ischar(msg) || isstring(msg)
            lines = cellstr(splitlines(string(msg)));
        else
            lines = {char(msg)};
        end
        cur = txtError.Value;
        if numel(cur)==1 && contains(cur{1},'(No messages yet)')
            cur = {};
        end
        txtError.Value = [cur; lines(:)];
        drawnow limitrate;
    end

    function reportError(ME)
        try
            rep = getReport(ME,'extended','hyperlinks','off');
        catch
            rep = ME.message;
        end
        txtError.Value = cellstr(splitlines(string(rep)));
    end

    % ---------------------------------------------------------------------
    % Export helper
    % ---------------------------------------------------------------------
    function outPath = exportFigureSmart(hFig, basePathNoExt, asPDF, forcePDF)
        if nargin < 4, forcePDF = false; end
        doPDF = logical(asPDF) || logical(forcePDF);

        if doPDF
            outPath = char(string(basePathNoExt) + ".pdf");

            wState = warning;
            cWarn = onCleanup(@() warning(wState)); 
            warning('off','all');
            try
                exportgraphics(hFig, outPath, 'ContentType','vector');
            catch
                set(hFig,'PaperPositionMode','auto');
                print(hFig, outPath, '-dpdf', '-vector');
            end
        else
            outPath = char(string(basePathNoExt) + ".jpg");
            exportgraphics(hFig, outPath, 'Resolution',600);
        end
    end

    function queuePPTImage(imgPath, titleStr)
        if nargin < 2 || strlength(string(titleStr)) == 0
            [~, nm, ~] = fileparts(char(imgPath));
            titleStr = strrep(nm,'_',' ');
        end

        pptQueuePaths(end+1,1)  = string(imgPath);
        pptQueueTitles(end+1,1) = string(titleStr);
    end

    function addSlideWithImage(ppt, imgPath, titleStr)
        slide = add(ppt,'Title and Content');

        if nargin < 3 || strlength(string(titleStr)) == 0
            [~, nm, ~] = fileparts(char(imgPath));
            titleStr = strrep(nm,'_',' ');
        else
            tS = string(titleStr);
            if contains(tS, filesep) || contains(tS, "/") || contains(tS, "\")
                [~, nm, ~] = fileparts(char(tS));
                titleStr = strrep(nm,'_',' ');
            end
        end

        replace(slide,'Title', char(string(titleStr)));
        replace(slide,'Content', mlreportgen.ppt.Picture(char(string(imgPath))));
    end

    % ---------------------------------------------------------------------
    % Plot helpers
    % ---------------------------------------------------------------------
    function addRowLabelTiled(ax, labelStr, xOffset, fsz)
        if nargin < 3 || isempty(xOffset), xOffset = -0.012; end
        if nargin < 4 || isempty(fsz), fsz = 11; end
        t = text(ax, xOffset, 0.5, char(string(labelStr)), ...
            'Units','normalized', ...
            'HorizontalAlignment','right', ...
            'VerticalAlignment','middle', ...
            'FontSize',fsz, ...
            'Rotation',0);
        t.Clipping = 'off';
    end

    function fillSegments(ax, maskArray, xVals, yBottom, yTop, fillColor, faceAlpha)
        if nargin < 7 || isempty(faceAlpha), faceAlpha = 0.3; end
        maskArray = logical(maskArray);
        diffMask  = diff([0 maskArray 0]);
        runStarts = find(diffMask==1);
        runEnds   = find(diffMask==-1)-1;
        for r = 1:numel(runStarts)
            idx = runStarts(r):runEnds(r);
            if isempty(idx), continue; end
            xPoly = [xVals(idx), fliplr(xVals(idx))];
            yPoly = [yBottom*ones(1,numel(idx)), yTop*ones(1,numel(idx))];
            fill(ax, xPoly, yPoly, fillColor, 'EdgeColor','none', 'FaceAlpha',faceAlpha);
        end
    end

    function modes = buildLabelModes(axisStyle, labsDay, labsDate)
        axisStyle = string(axisStyle);
        if axisStyle == "Days only"
            modes = struct('tag',"Days",'labels',{labsDay});
        elseif axisStyle == "Dated only"
            modes = struct('tag',"Dates",'labels',{labsDate});
        else
            modes(1) = struct('tag',"Days",'labels',{labsDay});
            modes(2) = struct('tag',"Dates",'labels',{labsDate});
        end
    end

    function addRightSideLabel(hFig, labelText)
        try
            delete(findall(hFig,'Tag','RightSideBlanketLabelAxes'));
        catch
        end

        axO = axes('Parent',hFig, ...
            'Units','normalized', ...
            'Position',[0.992 0.20 0.008 0.60], ...
            'Visible','off', ...
            'HitTest','off');
        axO.Tag = 'RightSideBlanketLabelAxes';

        t = text(axO, 0.5, 0.5, labelText, ...
            'Units','normalized', ...
            'HorizontalAlignment','center', ...
            'VerticalAlignment','middle', ...
            'Rotation',90, ...
            'FontSize',12, ...
            'Clipping','off');
        t.HitTest = 'off';
    end

    function lbl = formatTickLabels(vals, fmt)
        if nargin < 2 || strlength(string(fmt)) == 0
            fmt = '%.1f';
        end
        vals = vals(:);
        lbl = cell(numel(vals),1);
        for k = 1:numel(vals)
            lbl{k} = sprintf(fmt, vals(k));
        end
    end

    function [tMinTick, tMaxTick] = outwardRoundTempTicks(tMinRaw, tMaxRaw, step)
        if nargin < 3 || isempty(step) || ~isfinite(step) || step <= 0
            step = 0.1;
        end
        if ~isfinite(tMinRaw) || ~isfinite(tMaxRaw)
            tMinTick = NaN; tMaxTick = NaN; return;
        end
        if tMinRaw > tMaxRaw
            tmp = tMinRaw; tMinRaw = tMaxRaw; tMaxRaw = tmp;
        end

        tMinTick = floor(tMinRaw/step)*step;
        tMaxTick = ceil(tMaxRaw/step)*step;

        if tMinTick == tMaxTick
            tMaxTick = tMinTick + step;
        end

        if abs(tMinTick) < step/2, tMinTick = 0; end
        if abs(tMaxTick) < step/2, tMaxTick = 0; end
    end

    function setClimSafe(ax, lims)
        try
            ax.CLim = lims;
        catch
            set(ax,'CLim',lims);
        end
    end

    function cm = blueAmberRedMap(n)
        if nargin < 1 || isempty(n), n = 256; end
        c1 = [0.10 0.25 0.85];
        c2 = [1.00 0.80 0.20];
        c3 = [0.80 0.10 0.10];
        x  = [0 0.65 1];
        xi = linspace(0,1,n)';
        cm = interp1(x, [c1; c2; c3], xi, 'pchip');
        cm = max(0,min(1,cm));
    end

    % ---------------------------------------------------------------------
    % Data loading and parsing
    % ---------------------------------------------------------------------
    function [T, metaOut, usedSheetOut] = loadRawData(xlsxPath)
        try
            sheets = sheetnames(xlsxPath);
        catch
            [~, sheets] = xlsfinfo(xlsxPath);
            sheets = string(sheets);
        end

        if any(strcmpi(sheets,'RawData'))
            usedSheetOut = sheets{find(strcmpi(sheets,'RawData'),1)};
        else
            usedSheetOut = sheets{1};
        end

        T = readtable(xlsxPath, 'Sheet', usedSheetOut, 'VariableNamingRule','preserve');

        metaOut.timeVar     = requireVarLoose(T, {'Time stamp','Timestamp','Time Stamp','Time'});
        metaOut.activityVar = requireVarLoose(T, {'Sum of vector (SVMg)','SVMg','SVM','Activity','Activity (SVMg)'});
        metaOut.lightVar    = requireVarLoose(T, {'Light level (LUX)','Light level (Lux)','Light (LUX)','Lux','Light'});

        tempVar = findVarLoose(T, {'Temperature','Temp','Temperature (C)','Temperature (°C)'});
        if tempVar == ""
            metaOut.hasTemp = false;
            metaOut.tempVar = "";
        else
            metaOut.hasTemp = true;
            metaOut.tempVar = tempVar;
        end
    end

    function v = requireVarLoose(T, candidates)
        v = findVarLoose(T, candidates);
        if v == ""
            vars = string(T.Properties.VariableNames);
            error('Missing required column. Tried: %s. Found columns: %s', ...
                strjoin(string(candidates), ', '), strjoin(vars, ', '));
        end
    end

    function v = findVarLoose(T, candidates)
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
    end

    function n = normaliseNames(s)
        s = lower(string(s));
        s = strip(s);
        s = regexprep(s, '\s+', ' ');
        n = s;
    end

    function ts = parseTimestampsSoft(tsRaw)
        if isdatetime(tsRaw)
            ts = tsRaw;
            ts.TimeZone = '';
            return;
        end

        if isnumeric(tsRaw)
            try
                ts = datetime(tsRaw, 'ConvertFrom','excel');
                ts.TimeZone = '';
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
                tTry.TimeZone = '';
                ok = ~isnat(tTry) & isnat(ts);
                ts(ok) = tTry(ok);
            catch
            end
            if all(~isnat(ts)), break; end
        end
    end

    % ---------------------------------------------------------------------
    % Epoch + metric helpers
    % ---------------------------------------------------------------------
    function e = snapEpochToDayDivisor(epochSecIn)
        cands = [60, 90, 120, 180, 300, 360, 600, 720, 900, 1200, 1800, 3600];
        [~, idx] = min(abs(cands - epochSecIn));
        e = cands(idx);
        if mod(86400, e) ~= 0
            e = 60;
        end
    end

    function m = slidingMeanNan(x, win, minFrac)
        x = double(x(:))';
        n = numel(x);
        if win > n
            m = nan(1, 1);
            return;
        end

        valid = ~isnan(x);
        x0 = x; x0(~valid) = 0;

        kernel = ones(1,win);
        sumX   = conv(x0, kernel, 'valid');
        cntX   = conv(double(valid), kernel, 'valid');

        minCnt = ceil(minFrac * win);
        m = sumX ./ cntX;
        m(cntX < minCnt) = nan;
    end

    function [ISv, IVv] = computeISIV_hourlyFromDaily(dailyMat, intervalMinsIn)
        [~, nBins] = size(dailyMat);

        timeAxisLocal = (0:nBins-1) * intervalMinsIn;
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
            ISv = nan;
            IVv = nan;
            return;
        end

        xValid = x(valid);
        grandMean = mean(xValid,'omitnan');

        meanByHour = mean(hourlyMat, 1, 'omitnan');
        p = 24;

        denom = sum((xValid - grandMean).^2, 'omitnan');
        numer = p * sum((meanByHour - grandMean).^2, 'omitnan');

        if denom <= 0 || isnan(denom)
            ISv = nan;
        else
            ISv = numer / denom;
        end

        validPairs = ~isnan(x(1:end-1)) & ~isnan(x(2:end));
        if nnz(validPairs) < 24
            IVv = nan;
            return;
        end

        dx = diff(x);
        mssd = mean(dx(validPairs).^2, 'omitnan');
        varx = mean((xValid - grandMean).^2, 'omitnan');

        if varx <= 0 || isnan(varx)
            IVv = nan;
        else
            IVv = mssd / varx;
        end
    end

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

end