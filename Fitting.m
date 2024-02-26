clear all; close all; clc;
format longE
formatSpec = '%.15f';
M = 150; N = 50;
x0 = 0; y0 = 0; z0 = 0;
tic
%% Params need to confirm
L1_start = 2; L1_interval = 0.1; L1_end = 9; %position of First referactor 
L2_start = 5; L2_interval = 0.1; L2_end = 10;%position of Second referactor 
a1 = 20;an = 25;d = 1;
%% ���
for w = 25  %: d : an
    for L1 = L1_start %: L1_interval : L1_end
        for L2 = L2_start %: L2_interval : L2_end
            sheet = round((L1 - L1_start) / L1_interval) + 1;
            pdx = round((L2 - L2_start) / L2_interval) + 1;
            filename_X0 = ['F:\FreeFormAssets\CurvedSurface\X0_', num2str(w), '.xlsx'];
            filename_Y0 = ['F:\FreeFormAssets\CurvedSurface\Y0_', num2str(w), '.xlsx'];
            filename_Z0 = ['F:\FreeFormAssets\CurvedSurface\Z0_', num2str(w), '.xlsx'];
            filename_X3 = ['F:\FreeFormAssets\CurvedSurface\X3_', num2str(w), '.xlsx'];
            filename_Y3 = ['F:\FreeFormAssets\CurvedSurface\Y3_', num2str(w), '.xlsx'];
            filename_Z3 = ['F:\FreeFormAssets\CurvedSurface\Z3_', num2str(w), '.xlsx'];
            filename_xx1 = ['F:\FreeFormAssets\CurvedSurface\xx1_', num2str(w), '.xlsx'];
            filename_yy1 = ['F:\FreeFormAssets\CurvedSurface\yy1_', num2str(w), '.xlsx'];
            filename_zz1 = ['F:\FreeFormAssets\CurvedSurface\zz1_', num2str(w), '.xlsx'];
            filename_xx2 = ['F:\FreeFormAssets\CurvedSurface\xx2_', num2str(w), '.xlsx'];
            filename_yy2 = ['F:\FreeFormAssets\CurvedSurface\yy2_', num2str(w), '.xlsx'];
            filename_zz2 = ['F:\FreeFormAssets\CurvedSurface\zz2_', num2str(w), '.xlsx'];
            xlRange = ['A' num2str(1 + (pdx - 1) * (M + 1)) ':EU' num2str(1 + (pdx - 1) * (M + 1) + (M - 1))];
            X0 = xlsread(filename_X0, sheet, xlRange);
            Y0 = xlsread(filename_Y0, sheet, xlRange);
            Z0 = xlsread(filename_Z0, sheet, xlRange); %虚拟面的网格划分坐标
            X3 = xlsread(filename_X3, sheet, xlRange);
            Y3 = xlsread(filename_Y3, sheet, xlRange);
            Z3 = xlsread(filename_Z3, sheet, xlRange); %出射光斑的网格划分坐标
            x1 = xlsread(filename_xx1, sheet, xlRange);
            y1 = xlsread(filename_yy1, sheet, xlRange);
            z1 = xlsread(filename_zz1, sheet, xlRange);
            x2 = xlsread(filename_xx2, sheet, xlRange);
            y2 = xlsread(filename_yy2, sheet, xlRange);
            z2 = xlsread(filename_zz2, sheet, xlRange);
            figure
            axis equal
            cdata = cat(3, zeros(size(X0)), zeros(size(X0)), zeros(size(X0))); %作用是使右边的出射面变成黑色
            mesh(X3, Y3, Z3, cdata); hold on %出射面
            plot3(X3, Y3, Z3, 'o', 'markersize', 4, 'markerfacecolor', 'k', 'markeredgecolor', 'k'); hold on
            plot3(X0, Y0, Z0, 'o', 'markersize', 4, 'markerfacecolor', 'k', 'markeredgecolor', 'k'); hold on
            xlabel('x'); ylabel('y'); zlabel('z');
            %拟合曲面求系数
            xx1 = reshape(x1, [7701, 1]); yy1 = reshape(y1, [7701, 1]); zz1 = reshape(z1, [7701, 1]);
            xx2 = reshape(x2, [7701, 1]); yy2 = reshape(y2, [7701, 1]); zz2 = reshape(z2, [7701, 1]);
            f_SM = fit([xx1, yy1], zz1, 'poly55'); % plot(f_SM, [xx1, yy1], zz1);
            p = coeffvalues(f_SM);
            f_FM = fit([xx2, yy2], zz2, 'poly55'); % plot(f_FM, [xx2, yy2], zz2);
            q = coeffvalues(f_FM);
            warning('off', 'all')

            %带入拟合面 查看拟合结果
            N = 512;
            [xx1, yy1] = meshgrid(linspace(min(min(x1)), max(max(x1)), N), linspace(min(min(y1)), max(max(y1)), N));
            zz1 = p(1) + p(2) .* xx1 + p(3) .* yy1 + p(4) .* xx1 .^ 2 + p(5) .* xx1 .* yy1 + p(6) .* yy1 .^ 2 + p(7) .* xx1 .^ 3 + p(8) .* xx1 .^ 2 .* yy1 + ...
                p(9) .* xx1 .* yy1 .^ 2 + p(10) .* yy1 .^ 3 + p(11) .* xx1 .^ 4 + p(12) .* xx1 .^ 3 .* yy1 + p(13) .* xx1 .^ 2 .* yy1 .^ 2 + ...
                p(14) .* xx1 .* yy1 .^ 3 + p(15) .* yy1 .^ 4 + p(16) .* xx1 .^ 5 + p(17) .* xx1 .^ 4 .* yy1 + p(18) .* xx1 .^ 3 .* yy1 .^ 2 + ...
                p(19) .* xx1 .^ 2 .* yy1 .^ 3 + p(20) .* xx1 .* yy1 .^ 4 + p(21) .* yy1 .^ 5;
            [xx2, yy2] = meshgrid(linspace(min(min(x2)), max(max(x2)), N), linspace(min(min(y2)), max(max(y2)), N));
            zz2 = q(1) + q(2) .* xx2 + q(3) .* yy2 + q(4) .* xx2 .^ 2 + q(5) .* xx2 .* yy2 + q(6) .* yy2 .^ 2 + q(7) .* xx2 .^ 3 + q(8) .* xx2 .^ 2 .* yy2 + ...
                q(9) .* xx2 .* yy2 .^ 2 + q(10) .* yy2 .^ 3 + q(11) .* xx2 .^ 4 + q(12) .* xx2 .^ 3 .* yy2 + q(13) .* xx2 .^ 2 .* yy2 .^ 2 + ...
                q(14) .* xx2 .* yy2 .^ 3 + q(15) .* yy2 .^ 4 + q(16) .* xx2 .^ 5 + q(17) .* xx2 .^ 4 .* yy2 + q(18) .* xx2 .^ 3 .* yy2 .^ 2 + ...
                q(19) .* xx2 .^ 2 .* yy2 .^ 3 + q(20) .* xx2 .* yy2 .^ 4 + q(21) .* yy2 .^ 5;
            surf(xx1, yy1, zz1); hold on
            surf(xx2, yy2, zz2);
            colormap('copper'); shading interp; %view(0,0)
        end
    end
end