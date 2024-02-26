clear all; close all; clc;
M = 150; N = 50;
n1 = 1; n2 = 1.51; n3 = 1;
L0 = 1; dz = 0.3; L3 = 10;
x0 = 0; y0 = 0; z0 = 0; %光源的位置
faix = pi / 6; faiy = pi / 18;
lambda = 1.55 * 10 ^ (-3);
tic
%% 要确定的参量：
L1_start = 2; L1_interval = 0.1; L1_end = 9; %第一个透镜的位置参量
L2_start = 5; L2_interval = 0.1; L2_end = 10;%第二个透镜的位置参量
a1 = 20;an = 25;d = 1;%出射面光斑半径
%% 迭代计算曲面上的点
File_Name_X3 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'X3_L3=' num2str(L3) 'w=' num2str(a1) '-' num2str(d) '-' num2str(an) '.xlsx'];
File_Name_Y3 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Y3_L3=' num2str(L3) 'w=' num2str(a1) '-' num2str(d) '-' num2str(an) '.xlsx'];
File_Name_Z3 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Z3_L3=' num2str(L3) 'w=' num2str(a1) '-' num2str(d) '-' num2str(an) '.xlsx'];
File_Name_X0 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'X0_L0=' num2str(L0) '.xlsx'];
File_Name_Y0 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Y0_L0=' num2str(L0) '.xlsx'];
File_Name_Z0 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Z0_L0=' num2str(L0) '.xlsx'];
for w = 25  %: d : an
    sheet1 = (w - a1) / d + 1;
    X0 = xlsread(File_Name_X0);
    Y0 = xlsread(File_Name_Y0);
    Z0 = xlsread(File_Name_Z0); %虚拟面的网格划分坐标
    X3 = xlsread(File_Name_X3, sheet1);
    Y3 = xlsread(File_Name_Y3, sheet1);
    Z3 = xlsread(File_Name_Z3, sheet1); %出射光斑的网格划分坐标
    X3 = X3 ./ 80;
    Z3 = Z3 ./ 80;
    figure
    mesh(X0, Y0, Z0)
    axis equal;
    for L1 = L1_start %: L1_interval : L1_end
        for L2 = L2_start %: L2_interval : L2_end
            figure
            l = linspace(0, L3);
            plot3(zeros(size(l)), l, zeros(size(l)), 'k', 'linewidth', 1, 'linestyle', '--'); hold on;
            cdata = cat(3, zeros(size(X0)), zeros(size(X0)), zeros(size(X0))); %作用是使右边的出射面变成黑色
            axis equal; grid on;
            xlabel('x'); ylabel('y'); zlabel('z');
            mesh(X0, Y0, Z0, cdata); hold on
            mesh(X3, Y3, Z3, cdata); hold on
            A3 = [0 1 0];
            for i = 1 : M - 1
                for j = 1 : N
                    
                    x1(1, j) = 0; y1(1, j) = L1; z1(1, j) = 0;
                    x2(1, j) = 0; y2(1, j) = L2; z2(1, j) = 0;
                    A1(1) = X0(i, j) - x0; A1(2) = Y0(i, j) - (y0 + dz * abs(sin(atan(Z0(i, j) / X0(i, j)))));A1(3) = Z0(i, j) - z0; A1 = normr(A1);
                    A2(1) = x2(i, j) - x1(i, j); A2(2) = y2(i, j) - y1(i, j); A2(3) = z2(i, j) - z1(i, j); A2 = normr(A2);
                    
                    syms N2x N2y N2z
                    N2 = [N2x, N2y, N2z]; [N2x, N2y, N2z] = vpasolve(n3 * A3 - n2 * A2 - (n3 * dot(N2, A3) - n2 * dot(N2, A2)) .* N2);
                    N2 = [N2x, N2y, N2z];
                    syms k2 dx2 dy2 dz2
                    k2 = (dy2 - Y3(i + 1, j)) / A3(2);
                    dx2 = k2 * A3(1) + X3(i + 1, j);
                    dz2 = k2 * A3(3) + Z3(i + 1, j);
                    
                    [dy2] = solve(N2x * (dx2 - x2(i, j)) + N2y * (dy2 - y2(i, j)) + N2z * (dz2 - z2(i, j)));
                    y2(i + 1, j) = double(dy2);
                    k2 = (y2(i + 1, j) - Y3(i + 1, j)) / A3(2);
                    x2(i + 1, j) = k2 * A3(1) + X3(i + 1, j);
                    z2(i + 1, j) = k2 * A3(3) + Z3(i + 1, j);

                    syms N1x N1y N1z
                    N1 = [N1x, N1y, N1z]; [N1x, N1y, N1z] = vpasolve(n2 * A2 - n1 * A1 - (n2 * dot(N1, A2) - n1 * dot(N1, A1)) .* N1);
                    N1 = [N1x, N1y, N1z];
                    A12(1) = X0(i + 1, j) - x0; A12(2) = Y0(i + 1, j) - (y0 + dz * abs(sin(atan(Z0(i + 1, j) / X0(i + 1, j))))); A12(3) = Z0(i + 1, j) - z0; A12 = normr(A12);
                    syms k1 dx1 dy1 dz1
                    k1 = (dx1 - x0) / A12(1);
                    dy1 = k1 * A12(2) + y0 + dz * abs(sin(atan(Z0(i + 1, j) / X0(i + 1, j))));
                    dz1 = k1 * A12(3) + z0;
                    
                    [dx1] = solve(N1x * (dx1 - x1(i, j)) + N1y * (dy1 - y1(i, j)) + N1z * (dz1 - z1(i, j)));
                    x1(i + 1, j) = double(dx1);
                    k1 = (x1(i + 1, j) - x0) / A12(1);
                    y1(i + 1, j) = k1 * A12(2) + y0 + dz * abs(sin(atan(Z0(i + 1, j) / X0(i + 1, j))));
                    z1(i + 1, j) = k1 * A12(3) + z0;
                end
            end
            
            for i = 1 : M
                for j = 1 : N
                    line([x0, x1(i, j)], [y0 + dz * abs(sin(atan(Z0(i, j) / X0(i, j)))), y1(i, j)], [z0, z1(i, j)], 'color', 'r'); hold on
                    line([x1(i, j), x2(i, j)], [y1(i, j), y2(i, j)], [z1(i, j), z2(i, j)], 'color', 'r'); hold on
                    line([x2(i, j), X3(i, j)], [y2(i, j), Y3(i, j)], [z2(i, j), Z3(i, j)], 'color', 'r'); hold on
                end
            end
            
            surf(x1, y1, z1); hold on;
            surf(x2, y2, z2); hold on;
            axis equal;
            set(gcf, 'numbertitle', 'off', 'Name', ['理想情况下w=' num2str(w), 'L1=' num2str(L1), 'L2=' num2str(L2), 'L3=' num2str(L3)])
            saveas(gcf, ['F:\FreeFormAssets\CurvedSurface\理想情况下w=' num2str(w), 'L1=' num2str(L1), 'L2=' num2str(L2), 'L3=' num2str(L3) '.png'])
            sheet = round((L1 - L1_start) / L1_interval) + 1; %外部循环????，存入一个sheet
            pdx = round((L2 - L2_start) / L2_interval) + 1; %记录内部循环次数
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
            xlswrite(filename_X0, X0, sheet, xlRange); %保存文件
            xlswrite(filename_Y0, Y0, sheet, xlRange); %保存文件
            xlswrite(filename_Z0, Z0, sheet, xlRange); %保存文件
            xlswrite(filename_X3, X3, sheet, xlRange); %保存文件
            xlswrite(filename_Y3, Y3, sheet, xlRange); %保存文件
            xlswrite(filename_Z3, Z3, sheet, xlRange); %保存文件
            xlswrite(filename_xx1, x1, sheet, xlRange); %保存文件
            xlswrite(filename_yy1, y1, sheet, xlRange); %保存文件
            xlswrite(filename_zz1, z1, sheet, xlRange); %保存文件
            xlswrite(filename_xx2, x2, sheet, xlRange); %保存文件
            xlswrite(filename_yy2, y2, sheet, xlRange); %保存文件
            xlswrite(filename_zz2, z2, sheet, xlRange); %保存文件
        end
    end
end
toc