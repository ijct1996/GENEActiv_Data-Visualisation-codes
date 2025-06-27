function Actograph_v7_forparticipants_GUI.m
% ------------------------------------------------------------------------
% Actograph Analysis & Summary GUI for Parkinson's Participants
% ------------------------------------------------------------------------
% A self-contained MATLAB app for processing wrist-worn actigraphy data.
% Load your raw data into the provided "GENEActiv Data Template.xlsx":
%   1. Open "GENEActiv Data Template.xlsx" (included with this code).
%   2. Copy your raw CSV output into the template's "RawData" sheet.
%   3. Save the template as your input file.
%
% This application provides a single, self-contained interface to:
%   • Select raw actigraphy data and an output folder  
%   • Enter a global light threshold (lux)  
%   • Choose to include only complete 24-hour days or all days  
%   • Choose axis style: anonymised Day #, dated dd/mm, or both  
%   • Select which outputs to generate (including combined profile)  
%   • Bin data into 1-minute resolution per day (auto-detected)  
%   • Generate weekly rest–activity profiles (bar plots)  
%   • Generate weekly daily activity bar charts (anonymised and/or dated)  
%   • Generate weekly low-activity call-outs (anonymised and/or dated)  
%   • Generate weekly activity heatmaps  
%   • Generate weekly daily temperature profiles  
%       – Bar plots with light shading  
%       – Fixed y-axis 15–50 °C, tick every 20 °C  
%   • Generate daily light tracker plots  
%       – Full-range and 0–10 LUX panels  
%   • Generate weekly light distribution area plots (week-specific)  
%   • Generate weekly combined profiles  
%       – Activity bars, light shading, temperature line on second y-axis  
%   • Compute L5 (lowest-5-hour) and M10 (highest-10-hour) activity metrics  
%   • Compute Interdaily Stability (IS) and Intradaily Variability (IV)  
%   • Export high-resolution JPEGs (white backgrounds)  
%   • Write an Excel workbook with three sheets:  
%       – "Summary": daily metrics, L5/M10, Min/Max Temperature  
%       – "Metrics": IS, IV, and normal ranges  
%       – "Definitions": glossary of terms and interpretation guidance  
%   • Compile all figures into a PowerPoint presentation  
%
% Quick Start
%   Run in MATLAB:  
%       >> Actograph_v7_forparticipants_GUI  
%   1. Ensure your data are in "GENEActiv Data Template.xlsx".  
%   2. Select that file and choose an output folder.  
%   3. Adjust light threshold, day filtering and axis style.  
%   4. Tick the outputs you need.  
%   5. Click "Run"—sit back, it'll do the rest.  
%
% Requirements
%   • MATLAB R2019a or newer  
%   • Report Generator toolbox (for PowerPoint export)  
%
% Notes
%   – Designed for Parkinson's cohort studies but easily adapted.  
%   – All figures saved as high-res JPEGs; Excel and PPT files are optional.  
%
% © 2025 [Isaiah J Ting, Lall Lab]  
% ------------------------------------------------------------------------

    clc; clearvars; close all;

    %----------------------------------------------------------------------
    % Enforce universal figure size and resolution
    %----------------------------------------------------------------------
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
    chkCompleteOnly = uicheckbox(fig,'Position',[20 460 260 22], ...
        'Text','Include only complete 24‐hour days');

    % Axis style dropdown
    uilabel(fig,'Position',[20 420 80 22],'Text','Axis style:');
    ddlAxisStyle = uidropdown(fig,'Position',[100 420 200 22], ...
        'Items',{'Days only','Dated only','Both'},'Value','Both');

    % Outputs selection panel
    pnlOutputs = uipanel(fig,'Title','Select Outputs','Position',[20 260 560 140]);
    chkAll             = uicheckbox(pnlOutputs,'Position',[10 100 100 20], ...
                          'Text','All outputs','Value',true,'ValueChangedFcn',@toggleAllOutputs);
    chkActivityProf    = uicheckbox(pnlOutputs,'Position',[10 75 200 20], ...
                          'Text','Activity profiles','Value',true);
    chkDailyActivity   = uicheckbox(pnlOutputs,'Position',[10 50 200 20], ...
                          'Text','Daily activity bar charts','Value',true);
    chkLowActivity     = uicheckbox(pnlOutputs,'Position',[10 25 200 20], ...
                          'Text','Low‐activity call‐outs','Value',true);
    chkHeatmap         = uicheckbox(pnlOutputs,'Position',[220 75 200 20], ...
                          'Text','Activity heatmaps','Value',true);
    chkTempProfiles    = uicheckbox(pnlOutputs,'Position',[220 50 200 20], ...
                          'Text','Temperature profiles','Value',true);
    chkCombined        = uicheckbox(pnlOutputs,'Position',[220 25 300 20], ...
                          'Text','Combined profile (Activity, Light, Temp)','Value',false);
    chkLightTracker    = uicheckbox(pnlOutputs,'Position',[420 75 200 20], ...
                          'Text','Daily light tracker','Value',true);
    chkLightDist       = uicheckbox(pnlOutputs,'Position',[420 50 200 20], ...
                          'Text','Light distribution','Value',true);
    chkExcelSummary    = uicheckbox(pnlOutputs,'Position',[10 0 200 20], ...
                          'Text','Excel summary','Value',true);
    chkExcelMetrics    = uicheckbox(pnlOutputs,'Position',[220 0 200 20], ...
                          'Text','Excel metrics','Value',true);
    chkPowerPoint      = uicheckbox(pnlOutputs,'Position',[420 0 200 20], ...
                          'Text','PowerPoint','Value',true);

    % Output folder selection
    uilabel(fig,'Position',[20 220 90 22],'Text','Output folder:');
    txtOutputFolder = uieditfield(fig,'text','Position',[100 220 340 22]);
    uibutton(fig,'push','Position',[450 220 120 22],'Text','Browse...', ...
        'ButtonPushedFcn',@selectOutputFolder);

    % Run button
    btnRun = uibutton(fig,'push','Position',[240 170 120 30],'Text','Run', ...
        'ButtonPushedFcn',@runAnalysis);

    % Status label
    lblStatus = uilabel(fig,'Position',[20 130 560 22],'Text','Ready', ...
                        'HorizontalAlignment','left');

    %% Callback: select input file
    function selectInputFile(~,~)
        [file,path] = uigetfile('*.xlsx','Select Excel file');
        if isequal(file,0), return; end
        txtInputFile.Value = fullfile(path,file);
    end

    %% Callback: select output folder
    function selectOutputFolder(~,~)
        folder = uigetdir('','Select output folder');
        if isequal(folder,0), return; end
        txtOutputFolder.Value = folder;
    end

    %% Callback: toggle all outputs
    function toggleAllOutputs(src,~)
        val = src.Value;
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

    %% Main analysis function
    function runAnalysis(~,~)
        try
            % Ensure at least one output is selected
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

            lblStatus.Text = 'Loading data...'; drawnow;

            % Validate input file
            inputFilePath = txtInputFile.Value;
            if isempty(inputFilePath) || ~isfile(inputFilePath)
                uialert(fig,'Please select a valid input file.','Input Error');
                return;
            end

            % Read raw data
            rawTable       = readtable(inputFilePath,'VariableNamingRule','preserve');
            activityCounts = rawTable.("Sum of vector (SVMg)");
            lightLevels    = rawTable.("Light level (LUX)");
            temperature    = rawTable.("Temperature");
            timestamps     = datetime(rawTable.("Time stamp"), ...
                                     'InputFormat','yyyy-MM-dd HH:mm:ss:SSS');

            % Global light threshold
            lightThreshold = numThreshold.Value;

            % Detect sampling interval
            lblStatus.Text = 'Detecting sampling interval...'; drawnow;
            timeDiffs      = diff(timestamps);
            medianInterval = median(timeDiffs);
            intervalMins   = minutes(medianInterval);
            binsPerDay     = round(1440 / intervalMins);

            % Bin into daily matrices
            lblStatus.Text = 'Binning data into daily matrices...'; drawnow;
            firstDayStart    = dateshift(timestamps(1),'start','day');
            elapsedMins      = minutes(timestamps - firstDayStart);
            totalDays        = ceil(days(timestamps(end) - firstDayStart));
            binEdges         = 0 : intervalMins : 1440 * totalDays;
            binIndices       = discretize(elapsedMins,binEdges);
            dayStartDates    = firstDayStart + days(0:(totalDays-1))';
            weekDayNames     = cellstr(datestr(dayStartDates,'dddd'));

            % Preallocate binned data
            binnedActivity    = nan(totalDays,binsPerDay);
            binnedLight       = nan(totalDays,binsPerDay);
            binnedTemperature = nan(totalDays,binsPerDay);

            % Fill binned matrices
            for dayIdx = 1:totalDays
                sel    = binIndices > (dayIdx-1)*binsPerDay & binIndices <= dayIdx*binsPerDay;
                relBin = binIndices(sel) - (dayIdx-1)*binsPerDay;
                binnedActivity(dayIdx,relBin)    = activityCounts(sel);
                binnedLight(dayIdx,relBin)       = lightLevels(sel);
                binnedTemperature(dayIdx,relBin) = temperature(sel);
            end

            % Filter to only complete days if requested
            if chkCompleteOnly.Value
                lblStatus.Text = 'Filtering to complete days only...'; drawnow;
                completeMask      = all(~isnan(binnedActivity),2);
                dayStartDates     = dayStartDates(completeMask);
                weekDayNames      = weekDayNames(completeMask);
                binnedActivity    = binnedActivity(completeMask,:);
                binnedLight       = binnedLight(completeMask,:);
                binnedTemperature = binnedTemperature(completeMask,:);
                totalDays         = sum(completeMask);
            end

            % Validate output folder
            outputFolderPath = txtOutputFolder.Value;
            if isempty(outputFolderPath) || ~isfolder(outputFolderPath)
                uialert(fig,'Please select a valid output folder.','Output Error');
                return;
            end

            % Prepare for plotting
            lblStatus.Text = 'Generating figures...'; drawnow;
            timeAxis      = (0:binsPerDay-1) * intervalMins;
            numberOfWeeks = ceil(totalDays / 7);

            % Vectorized daily metrics
            totalActivityPerDay = nansum(binnedActivity,2);
            hoursInLightPerDay  = sum(binnedLight > lightThreshold,2) / (60 / intervalMins);
            minTempPerDay       = min(binnedTemperature,[],2);
            maxTempPerDay       = max(binnedTemperature,[],2);

            % Compute L5/M10 and their start times
            L5Bins   = round(5*60/intervalMins);
            M10Bins  = round(10*60/intervalMins);
            L5Start  = NaT(totalDays,1);
            M10Start = NaT(totalDays,1);
            L5Mean   = nan(totalDays,1);
            M10Mean  = nan(totalDays,1);
            for d = 1:totalDays
                sig  = binnedActivity(d,:);
                cL5  = conv(sig, ones(1,L5Bins)/L5Bins, 'valid');
                cM10 = conv(sig, ones(1,M10Bins)/M10Bins, 'valid');
                [L5Mean(d), idx5 ] = min(cL5);
                [M10Mean(d),idx10] = max(cM10);
                L5Start(d)  = dayStartDates(d) + minutes((idx5-1)*intervalMins);
                M10Start(d) = dayStartDates(d) + minutes((idx10-1)*intervalMins);
            end
            L5TimeStrings  = string(timeofday(L5Start));
            M10TimeStrings = string(timeofday(M10Start));

            % Compute IS and IV
            dataMat    = binnedActivity;
            meanHourly = nanmean(dataMat,1);
            grandMean  = nanmean(dataMat(:));
            ISvalue    = (totalDays * nansum((meanHourly-grandMean).^2)) / nansum((dataMat(:)-grandMean).^2);
            flatData   = dataMat(:);
            vm         = ~isnan(flatData);
            diffSq     = diff(flatData(vm)).^2;
            IVvalue    = (sum(vm)*nansum(diffSq)/(sum(vm)-1)) / nansum((flatData(vm)-grandMean).^2);

            % Global temperature limits & ticks
            globalMinTemp = 15;
            globalMaxTemp = 50;
            tempTicks     = [globalMinTemp, globalMinTemp+20, globalMaxTemp];

            % Create subfolder for daily light tracker
            if chkLightTracker.Value
                lightTrackerFolder = fullfile(outputFolderPath,'DailyLightTracker');
                if ~isfolder(lightTrackerFolder)
                    mkdir(lightTrackerFolder);
                end
                % Daily light tracker plots
                for d = 1:totalDays
                    dateLabel = datestr(dayStartDates(d),'dd_mm_yyyy');
                    figLT = figure('Visible','off','Color','w','Position',[100 100 800 600]);

                    % Top: full range
                    ax1 = subplot(2,1,1);
                    plot(ax1, timeAxis, binnedLight(d,:), 'Color',[0 0.5 0.5],'LineWidth',1);
                    hold(ax1,'on');
                    yline(ax1, lightThreshold,':','Color',[0.5 0.5 0.5],'LineWidth',0.5);
                    hold(ax1,'off');
                    title(ax1, sprintf('Light Tracker — %s (Full Range)', ...
                          datestr(dayStartDates(d),'dd mmm yyyy')), ...
                          'FontSize',14,'HorizontalAlignment','center');
                    xlim(ax1,[0 1440]);
                    ax1.XTick = [];
                    set(ax1,'TickDir','out','Box','off','FontSize',11);
                    ylabel(ax1,'Light (LUX)','FontSize',12);

                    % Bottom: zoom 0–10
                    ax2 = subplot(2,1,2);
                    plot(ax2, timeAxis, binnedLight(d,:), 'Color',[0 0.5 0.5],'LineWidth',1);
                    hold(ax2,'on');
                    yline(ax2, lightThreshold,':','Color',[0.5 0.5 0.5],'LineWidth',0.5);
                    hold(ax2,'off');
                    title(ax2,'Light Tracker — 0–10 LUX','FontSize',14,'HorizontalAlignment','center');
                    xlim(ax2,[0 1440]);
                    ylim(ax2,[0 10]);
                    set(ax2,'TickDir','out','Box','off','FontSize',11);
                    xlabel(ax2,'Time of Day','FontSize',12);
                    ylabel(ax2,'Light (LUX)','FontSize',12);
                    xticks(ax2,0:240:1440);
                    xticklabels(ax2,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});

                    exportgraphics(figLT, fullfile(lightTrackerFolder, ...
                        sprintf('DailyLightTracker_%s.jpg',dateLabel)),'Resolution',600);
                    close(figLT);
                end
            end

            % Weekly loop for weekly figures
            for wk = 1:numberOfWeeks
                daysIdx    = (wk-1)*7 + (1:7);
                daysIdx    = daysIdx(daysIdx <= totalDays);
                nThisWeek  = numel(daysIdx);
                dateLabels = cellstr(datestr(dayStartDates(daysIdx),'dd/mm'));
                anyDays    = any(strcmp(ddlAxisStyle.Value,{'Days only','Both'}));
                anyDates   = any(strcmp(ddlAxisStyle.Value,{'Dated only','Both'}));

                % Week‐specific time mask for light distribution
                weekStart = dayStartDates(daysIdx(1));
                weekEnd   = dayStartDates(daysIdx(end)) + days(1);
                weekMask  = timestamps >= weekStart & timestamps < weekEnd;

                % 1) Activity profiles
                if chkActivityProf.Value
                    if anyDays
                        figAP = figure('Visible','off','Color','w');
                        for i = 1:nThisWeek
                            d = daysIdx(i);
                            ax = subplot(nThisWeek,1,i); hold(ax,'on');
                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;
                            mA = binnedActivity(d,:);
                            yMax = max(mA,[],'omitnan')*1.2;
                            fillSegments(ax,mD,timeAxis,0,yMax,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,0,yMax,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,mA,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                            ylim(ax,[0 yMax]);
                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);
                            xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            ylabel(ax,sprintf('Day %d',d),'FontSize',12);
                            set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
                            if i==1
                                title(ax,'Participant Activity Profile','FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figAP, fullfile(outputFolderPath, ...
                            sprintf('01_ActivityProfile_Week%d.jpg',wk)),'Resolution',600);
                        close(figAP);
                    end
                    if anyDates
                        figAPD = figure('Visible','off','Color','w');
                        for i = 1:nThisWeek
                            d = daysIdx(i);
                            ax = subplot(nThisWeek,1,i); hold(ax,'on');
                            mL = binnedLight(d,:) > lightThreshold;
                            mD = ~mL;
                            mA = binnedActivity(d,:);
                            yMax = max(mA,[],'omitnan')*1.2;
                            fillSegments(ax,mD,timeAxis,0,yMax,[0.9 0.9 0.9]);
                            fillSegments(ax,mL,timeAxis,0,yMax,[0.9290 0.6940 0.1250]);
                            bar(ax,timeAxis,mA,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                            ylim(ax,[0 yMax]);
                            xlim(ax,[0 1440]);
                            xticks(ax,0:240:1440);
                            xticklabels(ax,{'00:00','04:00','08:00','12:00','16:00','20:00','24:00'});
                            ylabel(ax,dateLabels{i},'FontSize',12);
                            set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
                            if i==1
                                title(ax,'Participant Activity Profile (dd/mm)','FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figAPD, fullfile(outputFolderPath, ...
                            sprintf('01_ActivityProfile_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figAPD);
                    end
                end

                % 2) Daily activity bar charts
                if chkDailyActivity.Value
                    if anyDays
                        figDA = figure('Visible','off','Color','w');
                        bar(1:nThisWeek,totalActivityPerDay(daysIdx),'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        xticks(1:nThisWeek);
                        xticklabels(arrayfun(@(x)sprintf('Day %d',x),daysIdx,'UniformOutput',false));
                        xlabel('Day','FontSize',14);
                        title(sprintf('Total Activity by Day (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figDA, fullfile(outputFolderPath, ...
                            sprintf('02_DailyActivity_Week%d.jpg',wk)),'Resolution',600);
                        close(figDA);
                    end
                    if anyDates
                        figDAD = figure('Visible','off','Color','w');
                        bar(1:nThisWeek,totalActivityPerDay(daysIdx),'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        xticks(1:nThisWeek);
                        xticklabels(dateLabels);
                        xlabel('Date','FontSize',14);
                        title(sprintf('Total Activity by Date (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figDAD, fullfile(outputFolderPath, ...
                            sprintf('02_DailyActivity_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figDAD);
                    end
                end

                % 3) Low‐activity call‐outs
                if chkLowActivity.Value
                    weekTotals = totalActivityPerDay(daysIdx);
                    lowThresh  = mean(weekTotals) - std(weekTotals);
                    lowIdxs    = find(weekTotals < lowThresh);
                    if anyDays
                        figLC = figure('Visible','off','Color','w');
                        bar(1:nThisWeek,weekTotals,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        hold on;
                        for idx = lowIdxs'
                            text(idx, weekTotals(idx)+0.05*max(weekTotals),'Low','Color','r','FontSize',12,'HorizontalAlignment','center');
                        end
                        hold off;
                        xticks(1:nThisWeek);
                        xticklabels(arrayfun(@(x)sprintf('Day %d',x),daysIdx,'UniformOutput',false));
                        xlabel('Day','FontSize',14);
                        title(sprintf('Low‐Activity Call‐outs (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figLC, fullfile(outputFolderPath, ...
                            sprintf('03_LowActivity_Week%d.jpg',wk)),'Resolution',600);
                        close(figLC);
                    end
                    if anyDates
                        figLCD = figure('Visible','off','Color','w');
                        bar(1:nThisWeek,weekTotals,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        hold on;
                        for idx = lowIdxs'
                            text(idx, weekTotals(idx)+0.05*max(weekTotals),'Low','Color','r','FontSize',12,'HorizontalAlignment','center');
                        end
                        hold off;
                        xticks(1:nThisWeek);
                        xticklabels(dateLabels);
                        xlabel('Date','FontSize',14);
                        title(sprintf('Low‐Activity Call‐outs by Date (Week %d)',wk),'FontSize',16);
                        ylabel('Total Activity','FontSize',14);
                        set(gca,'TickDir','out','FontSize',12,'Box','off');
                        exportgraphics(figLCD, fullfile(outputFolderPath, ...
                            sprintf('03_LowActivity_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figLCD);
                    end
                end

                % 4) Activity heatmaps
                if chkHeatmap.Value
                    blockAct = binnedActivity(daysIdx,:);
                    figHM = figure('Visible','off','Color','w');
                    imagesc(timeAxis,1:nThisWeek,blockAct);
                    axis xy; set(gca,'YDir','reverse');
                    cMin = prctile(blockAct(~isnan(blockAct)),5);
                    cMax = prctile(blockAct(~isnan(blockAct)),95);
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
                    exportgraphics(figHM, fullfile(outputFolderPath, ...
                        sprintf('04_ActivityHeatmap_Week%d.jpg',wk)),'Resolution',600);
                    close(figHM);
                end

                % 5) Temperature profiles
                if chkTempProfiles.Value
                    if anyDays
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
                            if i==1
                                title(ax,sprintf('Temperature Profile — Week %d',wk),'FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figTP, fullfile(outputFolderPath, ...
                            sprintf('05_TemperatureProfile_Week%d.jpg',wk)),'Resolution',600);
                        close(figTP);
                    end
                    if anyDates
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
                            if i==1
                                title(ax,sprintf('Temperature Profile — Week %d (dd/mm)',wk),'FontSize',16);
                            end
                            hold(ax,'off');
                        end
                        xlabel('Time of Day','FontSize',12);
                        exportgraphics(figTPD, fullfile(outputFolderPath, ...
                            sprintf('05_TemperatureProfile_Week%d_dated.jpg',wk)),'Resolution',600);
                        close(figTPD);
                    end
                end

                % 6) Weekly light distribution
                if chkLightDist.Value
                    startDateStr = datestr(dayStartDates(daysIdx(1)),'dd/mm');
                    endDateStr   = datestr(dayStartDates(daysIdx(end)),'dd/mm');
                    figLD = figure('Visible','off','Color','w');
                    hod   = hour(timestamps(weekMask));
                    dataL = lightLevels(weekMask);
                    hourlyAvgWeek = arrayfun(@(h) mean(dataL(hod==h),'omitnan'),0:23);
                    area(0:23,hourlyAvgWeek,'FaceColor',[0 0.5 0.5]);
                    title(sprintf('Light Distribution - Week %d (%s - %s)', wk, startDateStr, endDateStr), ...
                          'FontSize',16);
                    xlabel('Hour of Day (0 = Midnight)','FontSize',14);
                    ylabel('Average Light (LUX)','FontSize',14);
                    xlim([0 23]); xticks(0:1:23);
                    set(gca,'TickDir','out','FontSize',11,'Box','off');
                    exportgraphics(figLD, fullfile(outputFolderPath, ...
                        sprintf('06_LightDistribution_Week%d.jpg',wk)),'Resolution',600);
                    close(figLD);
                end

                % 7) Combined profile (Activity, Light, Temp) with overlaid axes
                if chkCombined.Value
                    figCP = figure('Visible','off','Color','w');
                    for i = 1:nThisWeek
                        d = daysIdx(i);
                
                        % Draw activity bars and light shading
                        axMain = subplot(nThisWeek,1,i);
                        hold(axMain,'on');
                        axMain.XAxisLocation = 'origin';             % anchor x-axis at y=0
                        axMain.XLim       = [0 1440];
                        axMain.XTick      = 0:240:1440;
                        axMain.XTickLabel = {'00:00','04:00','08:00','12:00','16:00','20:00','24:00'};
                
                        mL = binnedLight(d,:) > lightThreshold;
                        mD = ~mL;
                        yMax = max(binnedActivity(d,:),[],'omitnan') * 1.2;
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
                
                        % Overlay transparent temperature axes
                        pos    = axMain.Position;
                        axTemp = axes('Position',pos, ...
                                      'Color','none', ...
                                      'YAxisLocation','right', ...
                                      'XAxisLocation','bottom', ...
                                      'Box','off', ...
                                      'TickDir','out', ...
                                      'FontSize',11, ...
                                      'XLim',axMain.XLim, ...
                                      'XTick',axMain.XTick);
                        hold(axTemp,'on');
                        axTemp.XTick = [];
                        axTemp.XAxis.Visible = 'off';
                
                        plot(axTemp,timeAxis,binnedTemperature(d,:),'LineWidth',1.5,'Color',[1 0 0]);
                        ylim(axTemp,[0 globalMaxTemp]);
                        yticks(axTemp,[0 tempTicks]);
                        ylabel(axTemp,'°C','FontSize',12);
                
                        linkaxes([axMain,axTemp],'x');
                
                        if i==1
                            title(axMain,sprintf('Combined Profile — Week %d',wk),'FontSize',16);
                        end
                
                        hold(axMain,'off');
                        hold(axTemp,'off');
                    end
                
                    xlabel(axMain,'Time of Day','FontSize',12);
                    exportgraphics(figCP, fullfile(outputFolderPath, ...
                                  sprintf('07_CombinedProfile_Week%d.jpg',wk)),'Resolution',600);
                    close(figCP);
                end

            end  % end weekly loop

            %------------------------------------------------------------------
            % Write Excel workbook
            %------------------------------------------------------------------
            summaryFile = fullfile(outputFolderPath,'10_Participant_Results.xlsx');

            if chkExcelSummary.Value
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
                MetricsTable = table( ...
                    ISvalue, ...
                    IVvalue, ...
                    "0.58–0.73", ...
                    "0.56–0.77", ...
                    'VariableNames',{'InterdailyStability','IntradailyVariability', ...
                                     'IS_NormalRange','IV_NormalRange'} );
                writetable(MetricsTable, summaryFile, 'Sheet','Metrics');
            end

            % Always write Definitions sheet
            definitions = {
              'Date','Calendar date of recording','Aligns metrics to calendar days';
              'Weekday','Day of the week','Distinguishes weekday versus weekend';
              'Day','Sequential day index','Day number since start';
              'TotalActivity','Sum of minute-by-minute activity','Overall daily movement';
              'HoursInLight','Total hours with LUX > threshold','Duration of light exposure';
              'L5_StartTime','Clock time when five-hour window of lowest mean begins','Rest-activity trough onset';
              'L5_Mean','Mean activity during that window','Depth of rest trough';
              'M10_StartTime','Clock time when ten-hour window of highest mean begins','Activity peak onset';
              'M10_Mean','Mean activity during that window','Strength of activity peak';
              'MinTemperature','Minimum recorded temperature per day (°C)','Daily minimum temperature';
              'MaxTemperature','Maximum recorded temperature per day (°C)','Daily maximum temperature'
            };
            writecell([{'Term','Definition','Interpretation'}; definitions], ...
                      summaryFile, 'Sheet','Definitions');

            %------------------------------------------------------------------
            % Compile PowerPoint
            %------------------------------------------------------------------
            if chkPowerPoint.Value
                lblStatus.Text = 'Creating PowerPoint...'; drawnow;
                import mlreportgen.ppt.*;
                pptFile = fullfile(outputFolderPath,'AllFigures_Report.pptx');
                ppt     = Presentation(pptFile);
                open(ppt);
                jpgFiles = dir(fullfile(outputFolderPath,'*.jpg'));
                for k = 1:numel(jpgFiles)
                    slide = add(ppt,'Title and Content');
                    replace(slide,'Title',erase(jpgFiles(k).name,'.jpg'));
                    replace(slide,'Content',Picture(fullfile(outputFolderPath,jpgFiles(k).name)));
                end
                close(ppt);
            end

            lblStatus.Text = 'Done.'; drawnow;
            pause(1);
            close(fig);

        catch ME
            uialert(fig,ME.message,'Error');
            lblStatus.Text = 'Error encountered.';
        end
    end

    %% Helper: fill contiguous segments of a logical mask
    function fillSegments(ax,maskArray,xVals,yBottom,yTop,fillColor)
        diffMask  = diff([0 maskArray 0]);
        runStarts = find(diffMask==1);
        runEnds   = find(diffMask==-1)-1;
        for r = 1:numel(runStarts)
            idx   = runStarts(r):runEnds(r);
            xPoly = [xVals(idx), fliplr(xVals(idx))];
            yPoly = [yBottom*ones(1,numel(idx)), yTop*ones(1,numel(idx))];
            fill(ax,xPoly,yPoly,fillColor,'EdgeColor','none','FaceAlpha',0.3);
        end
    end

end