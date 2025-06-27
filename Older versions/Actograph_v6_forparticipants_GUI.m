function ActographAnalysisApp
% ------------------------------------------------------------------------
% App: Actograph Analysis and Summary for Parkinson's Participants
% ------------------------------------------------------------------------
% This application provides a single, self‐contained interface to:
%   • Select raw actigraphy data and an output folder
%   • Choose to include only complete 24-hour days or all days
%   • Choose axis style: anonymised Day #, dated dd/mm, or both
%   • Bin data into 1-minute resolution per day
%   • Generate weekly rest–activity profiles (bar plots) and heatmaps
%   • Generate weekly daily summary bar charts of total activity and light exposure
%   • Highlight low-activity days with call-outs
%   • Produce a full-period light daily distribution area plot
%   • Compute L5 (lowest-5-hour) and M10 (highest-10-hour) activity metrics
%   • Compute Interdaily Stability (IS) and Intradaily Variability (IV)
%   • Export high-resolution JPEGs (with white backgrounds preserved)
%   • Write an Excel workbook with two sheets:
%       – "Summary": all daily metrics, L5/M10, IS/IV
%       – "Definitions": glossary of terms and interpretation guidance
%   • Compile all figures into a PowerPoint presentation
%
% Key UI features:
%   • File picker for the input `.xlsx` file
%   • Checkbox for filtering to complete days
%   • Dropdown for axis style selection
%   • Folder picker for outputs
%   • Run button with status updates displayed in the same window
%
% How to use:
%   1. Run this function in MATLAB.
%   2. In the window that appears:
%        – Browse for the input actigraphy file.
%        – Tick "Include only complete 24-hour days" if required.
%        – Select "Days only", "Dated only" or "Both" for axis style.
%        – Browse for or create an output folder.
%        – Click "Run" and monitor progress via the status label.
%   3. On completion, the GUI will close automatically.
% ------------------------------------------------------------------------

    % Create UI figure
    fig = uifigure('Name','Actograph Analysis','Position',[300 200 520 360]);

    % Input file selection
    uilabel(fig,'Position',[20 310 70 22],'Text','Input file:');
    txtFile = uieditfield(fig,'text','Position',[100 310 300 22],...
        'Tooltip','Full path to the input Excel file');
    uibutton(fig,'push','Position',[410 310 80 22],'Text','Browse...','Tooltip', ...
        'Select input Excel file','ButtonPushedFcn',@(~,~)selectFile());

    % Complete days only checkbox
    chkComplete = uicheckbox(fig,'Position',[100 275 240 22],...
        'Text','Include only complete 24‐hour days',...
        'Tooltip','Exclude any days lacking full 1440 minutes of data');

    % Axis style dropdown
    uilabel(fig,'Position',[20 240 80 22],'Text','Axis style:');
    ddlAxis = uidropdown(fig,'Position',[100 240 200 22],...
        'Items',{'Days only','Dated only','Both'},...
        'Value','Both',...
        'Tooltip','Select anonymised day numbers, dates, or both');

    % Output folder selection
    uilabel(fig,'Position',[20 205 80 22],'Text','Output folder:');
    txtOutput = uieditfield(fig,'text','Position',[100 205 300 22],...
        'Tooltip','Full path to output folder');
    uibutton(fig,'push','Position',[410 205 80 22],'Text','Browse...','Tooltip', ...
        'Select output folder','ButtonPushedFcn',@(~,~)selectOutput());

    % Run button
    btnRun = uibutton(fig,'push','Position',[200 150 120 30],...
        'Text','Run','Tooltip','Start analysis',...
        'ButtonPushedFcn',@(~,~)runAnalysis());

    % Status label
    lblStatus = uilabel(fig,'Position',[20 100 480 22],...
        'Text','Ready','HorizontalAlignment','left');

    % Callback: file browse
    function selectFile()
        [file,path] = uigetfile('*.xlsx','Select Excel file');
        if isequal(file,0), return; end
        txtFile.Value = fullfile(path,file);
    end

    % Callback: output folder browse
    function selectOutput()
        folder = uigetdir('','Select output folder');
        if isequal(folder,0), return; 
        end
        txtOutput.Value = folder;
    end

    % Main analysis callback
    function runAnalysis()
        try
            lblStatus.Text = 'Loading data...'; 
            drawnow;

            % Validate input file
            inputFile = txtFile.Value;
            if isempty(inputFile) || ~isfile(inputFile)
                uialert(fig,'Please select a valid input file.','Input Error');
                return
            end
            dataTable = readtable(inputFile,'VariableNamingRule','preserve');
            activityCounts = dataTable.("Sum of vector (SVMg)");
            lightLevels    = dataTable.("Light level (LUX)");
            buttonEvents   = dataTable.("Button (1/0)");
            timestamps     = datetime(dataTable.("Time stamp"),...
                                'InputFormat','yyyy-MM-dd HH:mm:ss:SSS');

            lblStatus.Text = 'Binning data into days...'; 
            drawnow;
            % Bin into days × minutes
            firstDayStart    = dateshift(timestamps(1),'start','day');
            elapsedMinutes   = minutes(timestamps - firstDayStart);
            totalDays        = ceil(days(timestamps(end)-firstDayStart));
            binEdges         = 0:1440*totalDays;
            minuteBinIndices = discretize(elapsedMinutes,binEdges);
            dayStartDates = firstDayStart + days(0:(totalDays-1))';
            weekDayNames  = cellstr(datestr(dayStartDates,'dddd'));
            binnedActivity = nan(totalDays,1440);
            binnedLight    = nan(totalDays,1440);
            binnedButton   = nan(totalDays,1440);
            for d=1:totalDays
                sel    = minuteBinIndices>(d-1)*1440 & minuteBinIndices<=d*1440;
                relMin = minuteBinIndices(sel)-(d-1)*1440;
                binnedActivity(d,relMin)=activityCounts(sel);
                binnedLight(d,relMin)   =lightLevels(sel);
                binnedButton(d,relMin)  =buttonEvents(sel);
            end

            % Filter complete days
            if chkComplete.Value
                lblStatus.Text = 'Filtering complete days...'; 
                drawnow;
                completeMask = all(~isnan(binnedActivity),2);
                dayStartDates  = dayStartDates(completeMask);
                weekDayNames   = weekDayNames(completeMask);
                binnedActivity = binnedActivity(completeMask,:);
                binnedLight    = binnedLight(completeMask,:);
                totalDays      = sum(completeMask);
            end

            % Validate output folder
            outputFolderPath = txtOutput.Value;
            if isempty(outputFolderPath) || ~isfolder(outputFolderPath)
                uialert(fig,'Please select a valid output folder.','Output Error');
                return
            end

            lblStatus.Text = 'Generating figures...'; 
            drawnow;
            timeAxis = linspace(0,1440,1440);
            nWeeks   = ceil(totalDays/7);

            for wk=1:nWeeks
                daysIdx   = (wk-1)*7+(1:7);
                daysIdx   = daysIdx(daysIdx<=totalDays);
                nThisWeek = numel(daysIdx);
                dateLabels = cellstr(datestr(dayStartDates(daysIdx),'dd/mm'));

                % Participant Activity Profile
                if any(strcmp(ddlAxis.Value,{'Days only','Both'}))
                    fig1 = figure('Visible','off','Color','w');
                    for i=1:nThisWeek
                        d=daysIdx(i); ax=subplot(nThisWeek,1,i); hold(ax,'on');
                        darkMask  = binnedLight(d,:)<=1;
                        lightMask = binnedLight(d,:)>1;
                        yMax = max(binnedActivity(d,:),[],'omitnan')*1.2;
                        if any(darkMask)
                            xD=[timeAxis(darkMask),fliplr(timeAxis(darkMask))];
                            yD=[zeros(1,sum(darkMask)),yMax*ones(1,sum(darkMask))];
                            fill(ax,xD,yD,[0.9 0.9 0.9],'EdgeColor','none','FaceAlpha',0.3);
                        end
                        if any(lightMask)
                            xL=[timeAxis(lightMask),fliplr(timeAxis(lightMask))];
                            yL=[zeros(1,sum(lightMask)),yMax*ones(1,sum(lightMask))];
                            fill(ax,xL,yL,[0.9290 0.6940 0.1250],'EdgeColor','none','FaceAlpha',0.3);
                        end
                        bar(ax,timeAxis,binnedActivity(d,:),'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        ylim(ax,[0 yMax]); 
                        xlim(ax,[0 1440]);
                        xticks(ax,0:120:1440);
                        xticklabels(ax,{'00:00','02:00','04:00','06:00','08:00','10:00', ...
                                       '12:00','14:00','16:00','18:00','20:00','22:00','00:00'});
                        ylabel(ax,sprintf('Day %d',d),'FontSize',12);
                        set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
                        if i==1 
                            title(ax,'Participant Activity Profile','FontSize',16); 
                        end
                        hold(ax,'off');
                    end
                    xlabel('Time of Day','FontSize',12);
                    exportgraphics(fig1, fullfile(outputFolderPath,sprintf('01_ActivityProfile_Week%d.jpg',wk)),'Resolution',600);
                    close(fig1);
                end
                if any(strcmp(ddlAxis.Value,{'Dated only','Both'}))
                    fig2 = figure('Visible','off','Color','w');
                    for i=1:nThisWeek
                        d=daysIdx(i); ax=subplot(nThisWeek,1,i); 
                        hold(ax,'on');
                        darkMask  = binnedLight(d,:)<=1;
                        lightMask = binnedLight(d,:)>1;
                        yMax = max(binnedActivity(d,:),[],'omitnan')*1.2;
                        if any(darkMask)
                            xD=[timeAxis(darkMask),fliplr(timeAxis(darkMask))];
                            yD=[zeros(1,sum(darkMask)),yMax*ones(1,sum(darkMask))];
                            fill(ax,xD,yD,[0.9 0.9 0.9],'EdgeColor','none','FaceAlpha',0.3);
                        end
                        if any(lightMask)
                            xL=[timeAxis(lightMask),fliplr(timeAxis(lightMask))];
                            yL=[zeros(1,sum(lightMask)),yMax*ones(1,sum(lightMask))];
                            fill(ax,xL,yL,[0.9290 0.6940 0.1250],'EdgeColor','none','FaceAlpha',0.3);
                        end
                        bar(ax,timeAxis,binnedActivity(d,:),'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                        ylim(ax,[0 yMax]); 
                        xlim(ax,[0 1440]);
                        xticks(ax,0:120:1440);
                        xticklabels(ax,{'00:00','02:00','04:00','06:00','08:00','10:00', ...
                                       '12:00','14:00','16:00','18:00','20:00','22:00','00:00'});
                        ylabel(ax,dateLabels{i},'FontSize',12);
                        set(ax,'TickDir','out','YTick',[],'Box','off','FontSize',11);
                        if i==1 
                            title(ax,'Participant Activity Profile (dd/mm)','FontSize',16); 
                        end
                        hold(ax,'off');
                    end
                    xlabel('Time of Day','FontSize',12);
                    exportgraphics(fig2, fullfile(outputFolderPath,sprintf('01_ActivityProfile_Week%d_dated.jpg',wk)),'Resolution',600);
                    close(fig2);
                end

                % Activity Heatmap
                blockData = binnedActivity(daysIdx,:);
                if any(strcmp(ddlAxis.Value,{'Days only','Both'}))
                    fig3 = figure('Visible','off','Color','w');
                    imagesc(timeAxis,1:nThisWeek,blockData); 
                    axis xy; 
                    set(gca,'YDir','reverse');
                    caxis([prctile(blockData(~isnan(blockData)),5),prctile(blockData(~isnan(blockData)),95)]);
                    colormap(parula); 
                    cb=colorbar; 
                    cb.Label.String='Activity (SVMg)';
                    xlabel('Time of Day','FontSize',12); 
                    ylabel('Day','FontSize',12);
                    xticks(0:360:1440); 
                    xticklabels({'00:00','06:00','12:00','18:00','24:00'});
                    set(gca,'TickDir','out','FontSize',11,'Box','off'); 
                    title('Activity Heatmap','FontSize',16);
                    exportgraphics(fig3, fullfile(outputFolderPath,sprintf('02_Activity_Heatmap_Week%d.jpg',wk)),'Resolution',600);
                    close(fig3);
                end
                if any(strcmp(ddlAxis.Value,{'Dated only','Both'}))
                    fig4 = figure('Visible','off','Color','w');
                    imagesc(timeAxis,1:nThisWeek,blockData); 
                    axis xy; 
                    set(gca,'YDir','reverse');
                    caxis([prctile(blockData(~isnan(blockData)),5),prctile(blockData(~isnan(blockData)),95)]);
                    colormap(parula); 
                    cb=colorbar; 
                    cb.Label.String='Activity (SVMg)';
                    xlabel('Time of Day','FontSize',12); 
                    yticks(1:nThisWeek); 
                    yticklabels(dateLabels);
                    xticks(0:360:1440); 
                    xticklabels({'00:00','06:00','12:00','18:00','24:00'});
                    set(gca,'TickDir','out','FontSize',11,'Box','off'); 
                    title('Activity Heatmap (dd/mm)','FontSize',16);
                    exportgraphics(fig4, fullfile(outputFolderPath,sprintf('02_Activity_Heatmap_Week%d_dated.jpg',wk)),'Resolution',600);
                    close(fig4);
                end

                % Daily Activity Bar Chart
                dailyTotals = nansum(binnedActivity(daysIdx,:),2);
                if any(strcmp(ddlAxis.Value,{'Days only','Both'}))
                    fig5 = figure('Visible','off','Color','w');
                    bar(1:nThisWeek,dailyTotals,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                    xticks(1:nThisWeek); 
                    xticklabels(arrayfun(@(d)sprintf('Day %d',d),daysIdx,'UniformOutput',false));
                    ylabel('Total Activity','FontSize',14); 
                    xlabel('Day','FontSize',14);
                    set(gca,'TickDir','out','FontSize',12,'Box','off'); 
                    title('Total Activity by Day','FontSize',16);
                    exportgraphics(fig5, fullfile(outputFolderPath,sprintf('03_DailyActivity_Week%d.jpg',wk)),'Resolution',600);
                    close(fig5);
                end
                if any(strcmp(ddlAxis.Value,{'Dated only','Both'}))
                    fig6 = figure('Visible','off','Color','w');
                    bar(1:nThisWeek,dailyTotals,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none');
                    xticks(1:nThisWeek); 
                    xticklabels(dateLabels);
                    ylabel('Total Activity','FontSize',14); 
                    xlabel('Date','FontSize',14);
                    set(gca,'TickDir','out','FontSize',12,'Box','off'); 
                    title('Total Activity by Date (dd/mm)','FontSize',16);
                    exportgraphics(fig6, fullfile(outputFolderPath,sprintf('03_DailyActivity_Week%d_dated.jpg',wk)),'Resolution',600);
                    close(fig6);
                end

                % Daily Light Exposure Bar Chart
                hoursLight = sum(binnedLight(daysIdx,:)>1,2)/60;
                if any(strcmp(ddlAxis.Value,{'Days only','Both'}))
                    fig7 = figure('Visible','off','Color','w');
                    bar(1:nThisWeek,hoursLight,'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');
                    xticks(1:nThisWeek); 
                    xticklabels(arrayfun(@(d)sprintf('Day %d',d),daysIdx,'UniformOutput',false));
                    ylabel('Hours in Light','FontSize',14); 
                    xlabel('Day','FontSize',14);
                    set(gca,'TickDir','out','FontSize',12,'Box','off'); 
                    title('Daily Hours in Light','FontSize',16);
                    exportgraphics(fig7, fullfile(outputFolderPath,sprintf('04_DailyLight_Week%d.jpg',wk)),'Resolution',600);
                    close(fig7);
                end
                if any(strcmp(ddlAxis.Value,{'Dated only','Both'}))
                    fig8 = figure('Visible','off','Color','w');
                    bar(1:nThisWeek,hoursLight,'FaceColor',[0.8500 0.3250 0.0980],'EdgeColor','none');
                    xticks(1:nThisWeek); 
                    xticklabels(dateLabels);
                    ylabel('Hours in Light','FontSize',14); 
                    xlabel('Date','FontSize',14);
                    set(gca,'TickDir','out','FontSize',12,'Box','off'); 
                    title('Daily Hours in Light by Date (dd/mm)','FontSize',16);
                    exportgraphics(fig8, fullfile(outputFolderPath,sprintf('04_DailyLight_Week%d_dated.jpg',wk)),'Resolution',600);
                    close(fig8);
                end

                % Low‐Activity Call‐outs
                lowThresh = mean(dailyTotals)-std(dailyTotals);
                lowIdxs   = find(dailyTotals<lowThresh);
                if any(strcmp(ddlAxis.Value,{'Days only','Both'}))
                    fig9=figure('Visible','off','Color','w');
                    bar(1:nThisWeek,dailyTotals,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none'); hold on;
                    for idx = lowIdxs'
                        text(idx,dailyTotals(idx)+0.05*max(dailyTotals),'Low', ...
                             'Color','r','FontSize',12,'HorizontalAlign','center');
                    end
                    xticks(1:nThisWeek); xticklabels(arrayfun(@(d)sprintf('Day %d',d),daysIdx,'UniformOutput',false));
                    ylabel('Total      Activity','FontSize',14); 
                    xlabel('Day','FontSize',14);
                    set(gca,'TickDir','out','FontSize',12,'Box','off'); 
                    title('Low-Activity Call-outs','FontSize',16);
                    exportgraphics(fig9, fullfile(outputFolderPath,sprintf('05_LowActivity_Week%d.jpg',wk)),'Resolution',600);
                    close(fig9);
                end
                if any(strcmp(ddlAxis.Value,{'Dated only','Both'}))
                    fig10=figure('Visible','off','Color','w');
                    bar(1:nThisWeek,dailyTotals,'FaceColor',[0 0.4470 0.7410],'EdgeColor','none'); 
                    hold on;
                    for idx = lowIdxs'
                        text(idx,dailyTotals(idx)+0.05*max(dailyTotals),'Low', ...
                             'Color','r','FontSize',12,'HorizontalAlign','center');
                    end
                    xticks(1:nThisWeek); 
                    xticklabels(dateLabels);
                    ylabel('Total      Activity','FontSize',14); 
                    xlabel('Date','FontSize',14);
                    set(gca,'TickDir','out','FontSize',12,'Box','off'); 
                    title('Low-Activity Call-outs by Date (dd/mm)','FontSize',16);
                    exportgraphics(fig10, fullfile(outputFolderPath,sprintf('05_LowActivity_Week%d_dated.jpg',wk)),'Resolution',600);
                    close(fig10);
                end
            end

            % Light Daily Distribution
            lblStatus.Text = 'Generating light distribution plot...'; 
            drawnow;
            hourOfDay     = hour(timestamps);
            hourlyAvg     = arrayfun(@(h)mean(lightLevels(hourOfDay==h),'omitnan'),0:23);
            figLD = figure('Visible','off','Color','w','Name','Light Daily Distribution', ...
                           'NumberTitle','off','Position',[100 100 900 500],'Toolbar','none');
            area(0:23,hourlyAvg,'FaceColor',[0.4 0.7608 0.6471]);
            title('Light Daily Distribution','FontSize',16,'FontWeight','bold');
            xlabel('Hour of Day (0 = Midnight)','FontSize',14);
            ylabel('Average Light (LUX)','FontSize',14);
            xlim([0 23]); 
            xticks(0:1:23);
            set(gca,'TickDir','out','FontSize',11,'Box','off');
            exportgraphics(figLD,fullfile(outputFolderPath,'06_LightDailyDistribution.jpg'),'Resolution',600);
            close(figLD);

            % Compute summary metrics and write Excel
            lblStatus.Text = 'Computing summary and writing Excel...'; 
            drawnow;

            % Daily indices
            daysAll = (1:totalDays)';
            totalActivity = nansum(binnedActivity,2);
            hoursInLight  = sum(binnedLight>1,2)/60;

            % Peak activity time
            peakIdxs  = arrayfun(@(d)find(binnedActivity(d,:)==max(binnedActivity(d,:)),1),daysAll);
            peakTime  = timeAxis(peakIdxs)';
            peakDur   = minutes(peakTime); peakDur.Format='hh:mm:ss';
            peakStr   = string(peakDur);

            % Light start/end
            lightStartMin = nan(totalDays,1); 
            lightEndMin = nan(totalDays,1);
            for d=1:totalDays
                mask= binnedLight(d,:)>1;
                if any(mask)
                    lightStartMin(d)=timeAxis(find(mask,1,'first'));
                    lightEndMin(d)=timeAxis(find(mask,1,'last'));
                end
            end
            startDur = minutes(lightStartMin); 
            startDur.Format='hh:mm:ss';
            endDur   = minutes(lightEndMin);   
            endDur.Format   ='hh:mm:ss';
            startStr = string(startDur); 
            endStr = string(endDur);

            % L5/M10 on activity
            L5W   =5*60; 
            M10W=10*60;
            L5start=NaT(totalDays,1); 
            M10start=NaT(totalDays,1);
            L5mean=nan(totalDays,1);    
            M10mean=nan(totalDays,1);
            for d=1:totalDays
                sig=binnedActivity(d,:);
                c5=conv(sig,ones(1,L5W)/L5W,'valid');
                c10=conv(sig,ones(1,M10W)/M10W,'valid');
                [L5mean(d),i5]=min(c5);
                [M10mean(d),i10]=max(c10);
                L5start(d)=dayStartDates(d)+minutes(i5-1);
                M10start(d)=dayStartDates(d)+minutes(i10-1);
            end
            L5str=string(timeofday(L5start));
            M10str=string(timeofday(M10start));

            % IS & IV
            Mmat = binnedActivity;
            mh   = nanmean(Mmat,1);
            gm   = nanmean(Mmat(:));
            IS   = (totalDays*nansum((mh-gm).^2)) / nansum((Mmat(:)-gm).^2);
            fd   = Mmat(:); v=~isnan(fd);
            d2   = diff(fd(v));
            IV   = (sum(v)*nansum(d2.^2)/(sum(v)-1)) / nansum((fd(v)-gm).^2);

            IScol=[IS;nan(totalDays-1,1)];
            IVcol=[IV;nan(totalDays-1,1)];
            ISnorm="0.58-0.73"; IVnorm="0.56-0.77";
            ISnormCol=repmat(ISnorm,totalDays,1);
            IVnormCol=repmat(IVnorm,totalDays,1);

            excelFile = fullfile(outputFolderPath,'07_Participant_results.xlsx');
            SummaryT = table(dayStartDates, weekDayNames, daysAll, totalActivity, hoursInLight, ...
                             peakStr, startStr, endStr, ...
                             L5str, L5mean, M10str, M10mean, ...
                             IScol, IVcol, ISnormCol, IVnormCol, ...
                             'VariableNames',{...
                              'Date','Weekday','Day','TotalActivity','HoursInLight',...
                              'PeakActivityTime','LightStartTime','LightEndTime',...
                              'L5_StartTime','L5_Mean','M10_StartTime','M10_Mean',...
                              'InterdailyStability','IntradailyVariability',...
                              'IS_NormalRange','IV_NormalRange'});
            writetable(SummaryT, excelFile, 'Sheet','Summary');

            % Definitions sheet
            defs = {
              'Date','Calendar date of recording','Aligns metrics to calendar days';
              'Weekday','Day of the week','Distinguishes weekday vs weekend';
              'Day','Sequential day index','Day number since start';
              'TotalActivity','Sum of minute‐by‐minute activity','Overall daily movement';
              'HoursInLight','Total hours with LUX >1','Duration of light exposure';
              'PeakActivityTime','Clock time of maximum activity','Indicates peak movement time';
              'LightStartTime','Clock time of first minute >1 LUX','Onset of daily light exposure';
              'LightEndTime','Clock time of last minute >1 LUX','End of daily light exposure';
              'L5_StartTime','Clock time when the five‐hour window of lowest mean activity begins','Rest‐activity trough onset';
              'L5_Mean','Average activity during that window','Depth of rest trough';
              'M10_StartTime','Clock time when the ten‐hour window of highest mean activity begins','Activity peak onset';
              'M10_Mean','Average activity during that window','Strength of activity peak';
              'InterdailyStability','Consistency of daily rhythm','Higher = more stable';
              'IntradailyVariability','Fragmentation within day','Higher = more fragmented';
              'IS_NormalRange','Typical IS range (0.58‐0.73)','Healthy adult reference';
              'IV_NormalRange','Typical IV range (0.56‐0.77)','Healthy adult reference'
            };
            writecell([{'Term','Definition','Interpretation'}; defs], excelFile, 'Sheet','Definitions');

            % PowerPoint bundling
            lblStatus.Text = 'Creating PowerPoint...'; 
            drawnow;
            import mlreportgen.ppt.*;
            pptFile = fullfile(outputFolderPath,'AllFigures_Report.pptx');
            ppt = Presentation(pptFile); 
            open(ppt);
            jpgs = dir(fullfile(outputFolderPath,'*.jpg'));
            for i=1:numel(jpgs)
                slide=add(ppt,'Title and Content');
                replace(slide,'Title',erase(jpgs(i).name,'.jpg'));
                pic=Picture(fullfile(outputFolderPath,jpgs(i).name));
                replace(slide,'Content',pic);
            end
            close(ppt);

            lblStatus.Text = 'Done.'; 
            drawnow;
            pause(1);    % let the user see "Done."
            close(fig);  % this closes the UIFigure window

        catch ME
            uialert(fig, ME.message, 'Error');
            lblStatus.Text = 'Error encountered.';
        end
    end
end
