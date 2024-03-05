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
%% ���?
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
            Z3 = xlsread(filename_Z3, sheet, xlRange); %出射光斑的网格划分坐�?
            x1 = xlsread(filename_xx1, sheet, xlRange);
            y1 = xlsread(filename_yy1, sheet, xlRange);
            z1 = xlsread(filename_zz1, sheet, xlRange);
            x2 = xlsread(filename_xx2, sheet, xlRange);
            y2 = xlsread(filename_yy2, sheet, xlRange);
            z2 = xlsread(filename_zz2, sheet, xlRange);
            figure
            axis equal
            cdata = cat(3, zeros(size(X0)), zeros(size(X0)), zeros(size(X0))); %作用是使右边的出射面变成黑色
            mesh(X3, Y3, Z3, cdata); hold on %出射�?
            plot3(X3, Y3, Z3, 'o', 'markersize', 4, 'markerfacecolor', 'k', 'markeredgecolor', 'k'); hold on
            plot3(X0, Y0, Z0, 'o', 'markersize', 4, 'markerfacecolor', 'k', 'markeredgecolor', 'k'); hold on
            xlabel('x'); ylabel('y'); zlabel('z');
            %拟合曲面求系�?(xx, yy, zz)代表的是拟合出来的平�?
            xx1 = reshape(x1, [M * N, 1]); yy1 = reshape(y1, [M * N, 1]); zz1 = reshape(z1, [M * N, 1]);
            xx2 = reshape(x2, [M * N, 1]); yy2 = reshape(y2, [M * N, 1]); zz2 = reshape(z2, [M * N, 1]);
            f_SM = fit([xx1, zz1], yy1, 'poly55'); %plot(f_SM, [xx1, yy1], zz1);
            p = coeffvalues(f_SM);
            f_FM = fit([xx2, zz2], yy2, 'poly55'); %plot(f_FM, [xx2, yy2], zz2);
            q = coeffvalues(f_FM);
            warning('off', 'all')

            %带入拟合�?查看拟合结果 (xxx, yyy, zzz)代表的是带入拟合面后的平�?
            T = 512;
            [xxx1, zzz1] = meshgrid(linspace(min(min(xx1)), max(max(xx1)), T), linspace(min(min(zz1)), max(max(zz1)), T));
            yyy1 = p(1) + p(2) .* xxx1 + p(3) .* zzz1 + p(4) .* xxx1 .^ 2 + p(5) .* xxx1 .* zzz1 + p(6) .* zzz1 .^ 2 + p(7) .* xxx1 .^ 3 + p(8) .* xxx1 .^ 2 .* zzz1 + ...
                p(9) .* xxx1 .* zzz1 .^ 2 + p(10) .* zzz1 .^ 3 + p(11) .* xxx1 .^ 4 + p(12) .* xxx1 .^ 3 .* zzz1 + p(13) .* xxx1 .^ 2 .* zzz1 .^ 2 + ...
                p(14) .* xxx1 .* zzz1 .^ 3 + p(15) .* zzz1 .^ 4 + p(16) .* xxx1 .^ 5 + p(17) .* xxx1 .^ 4 .* zzz1 + p(18) .* xxx1 .^ 3 .* zzz1 .^ 2 + ...
                p(19) .* xxx1 .^ 2 .* zzz1 .^ 3 + p(20) .* xxx1 .* zzz1 .^ 4 + p(21) .* zzz1 .^ 5;
            [xxx2, zzz2] = meshgrid(linspace(min(min(xx2)), max(max(xx2)), T), linspace(min(min(zz2)), max(max(zz2)), T));
            yyy2 = q(1) + q(2) .* xxx2 + q(3) .* zzz2 + q(4) .* xxx2 .^ 2 + q(5) .* xxx2 .* zzz2 + q(6) .* zzz2 .^ 2 + q(7) .* xxx2 .^ 3 + q(8) .* xxx2 .^ 2 .* zzz2 + ...
                q(9) .* xxx2 .* zzz2 .^ 2 + q(10) .* zzz2 .^ 3 + q(11) .* xxx2 .^ 4 + q(12) .* xxx2 .^ 3 .* zzz2 + q(13) .* xxx2 .^ 2 .* zzz2 .^ 2 + ...
                q(14) .* xxx2 .* zzz2 .^ 3 + q(15) .* zzz2 .^ 4 + q(16) .* xxx2 .^ 5 + q(17) .* xxx2 .^ 4 .* zzz2 + q(18) .* xxx2 .^ 3 .* zzz2 .^ 2 + ...
                q(19) .* xxx2 .^ 2 .* zzz2 .^ 3 + q(20) .* xxx2 .* zzz2 .^ 4 + q(21) .* zzz2 .^ 5;
            surf(xxx1, yyy1, zzz1); hold on
            surf(xxx2, yyy2, zzz2);
            colormap('copper'); shading interp; %view(0,0)
        end
    end
end