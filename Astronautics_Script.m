clc; clear; close all;

filename  = "Astronautics Data.xlsx";
sheetName = "Main Data";

envColsOriginal = ["Earth (J)", "Moon (J)", "Mars (J)", "Microgravity (J)"];
envLabels       = ["Earth", "Moon", "Mars", "Microgravity"];

T = readtable(filename, "Sheet", sheetName);

namesRaw = string(T{:,1});

origHeaders = string(T.Properties.VariableDescriptions);
if isempty(origHeaders) || all(origHeaders == "")
    origHeaders = string(T.Properties.VariableNames);
end

idx = zeros(1, numel(envColsOriginal));
for k = 1:numel(envColsOriginal)
    hit = find(origHeaders == envColsOriginal(k), 1);
    if isempty(hit)
        error("Could not find original header '%s'.\nAvailable headers:\n%s", ...
            envColsOriginal(k), strjoin(origHeaders, ", "));
    end
    idx(k) = hit;
end

Xraw = double(T{:, idx});

good  = all(isfinite(Xraw), 2) & (namesRaw ~= "");
names = namesRaw(good);
X     = Xraw(good, :);

if numel(names) ~= 5
    warning("Expected 5 group members, but found %d valid rows after filtering.", numel(names));
end

% Stats
mins  = min(X, [], 1);
maxs  = max(X, [], 1);
means = mean(X, 1);
stds  = std(X, 0, 1);

statsTable = table(envLabels(:), mins(:), maxs(:), means(:), stds(:), ...
    'VariableNames', ["Environment", "Min", "Max", "Mean", "StdDev"]);

disp("===== Summary Stats Across Group =====");
disp(statsTable);

%% ---- Visualization styling knobs (DOUBLED) ----
baseFont   = 32;   % axes + tick labels
labelFont  = 36;   % axis labels
titleFont  = 40;   % title
nameFont   = 20;   % individual name labels

dotSize    = 160;  % larger markers
meanLW     = 4.0;  % mean line width
whiskerLW  = 3.5;  % min/max line width
errLW      = 3.5;  % error bar line width
capSize    = 24;   % error bar cap size

%% ---- Plot ----
figure('Color','w'); hold on; grid on;

m = size(X,2);
n = size(X,1);

for j = 1:m
    xj = j * ones(n,1);
    yvals  = X(:,j);

    scatter(xj, yvals, dotSize, 'filled', 'MarkerFaceAlpha', 0.90);

    % Data-scaled vertical offsets (Option 2)
    yRange = max(yvals) - min(yvals);
    if yRange == 0
        yRange = 1;
    end
    labelOffset = linspace(-0.06*yRange, 0.06*yRange, n);

    for i = 1:n
        if mod(i,2) == 0
            % Right side
            xText = xj(i) + 0.10;
            hAlign = 'left';
        else
            % Left side
            xText = xj(i) - 0.10;
            hAlign = 'right';
        end

        text(xText, yvals(i) + labelOffset(i), names(i), ...
            'FontSize', nameFont, ...
            'HorizontalAlignment', hAlign, ...
            'VerticalAlignment', 'middle');
    end
end

% Mean ± std
errorbar(1:m, means, stds, 'k.', ...
    'LineWidth', errLW, 'CapSize', capSize);

% Min / max whiskers + mean bar
for j = 1:m
    plot([j j], [mins(j) maxs(j)], 'k-', 'LineWidth', whiskerLW);
    plot([j-0.20 j+0.20], [means(j) means(j)], 'k-', 'LineWidth', meanLW);
end

xlim([0.5, m+0.5]);
xticks(1:m);
xticklabels(envLabels);

ax = gca;
ax.FontSize  = baseFont;
ax.LineWidth = 2.5;

ylabel("Work / Load metric (J)", 'FontSize', labelFont);
title("Group Loads/Work by Environment", 'FontSize', titleFont);

hold off;

%% ---- Console summary ----
fprintf("\n===== Ranges (Min to Max) =====\n");
for j = 1:numel(envLabels)
    fprintf("%s: %.3f to %.3f (mean %.3f, std %.3f)\n", ...
        envLabels(j), mins(j), maxs(j), means(j), stds(j));
end
