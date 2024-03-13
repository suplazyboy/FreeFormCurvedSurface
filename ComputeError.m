clear all; close all; clc;
format longE
formatSpec = '%.15f';
M = 150, N = 50;dz = 0.3;
x0 = 0; y0 = 0; z0 = 0;
CMP = colormap(jet);
tic
%% Params need to confirm
L1_start = 2; L1_interval = 0.1; L1_end = 9; %position of First referactor 
L2_start = 5; L2_interval = 0.1; L2_end = 10;%position of Second referactor 
a1 = 20;an = 25;d = 1;

%% Compute Error
for w = 25  %: d : an
    for L1 = L1_start %: L1_interval : L1_end
        for L2 = L2_start %: L2_interval : L2_end
            sheet = round((L1 - L1_start) / L1_interval) + 1;
            pdx = round((L2 - L2_start) / L2_interval) + 1;
            xlRange = ['A' num2str(1 + (pdx - 1) * (M + 1)) ':EU' num2str(1 + (pdx - 1) * (M + 1) + (M - 1))];
            %读取理想情况下的点
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

            %拟合后的点
            dictionaryname = ['w=', num2str(w), ' L1=', num2str(L1), ' L2=', num2str(L2)];
            filename_X0 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_X0.xlsx'];
            filename_Y0 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_Y0.xlsx'];
            filename_Z0 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_Z0.xlsx'];
            filename_x3 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_x3.xlsx'];
            filename_y3 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_y3.xlsx'];
            filename_z3 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_z3.xlsx'];
            filename_X1 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_X1.xlsx'];
            filename_Y1 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_Y1.xlsx'];
            filename_Z1 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_Z1.xlsx'];
            filename_X2 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_X1.xlsx'];
            filename_Y2 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_Y2.xlsx'];
            filename_Z2 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'real_Z2.xlsx'];
            filename_alpha = ['F:\FreeFormAssets\RealPoint\', dictionaryname, 'alpha.xlsx']
            
            real_X0 = xlsread(filename_X0, sheet, xlRange);
            real_Y0 = xlsread(filename_Y0, sheet, xlRange);
            real_Z0 = xlsread(filename_Z0, sheet, xlRange);
            real_x3 = xlsread(filename_x3, sheet, xlRange);
            real_y3 = xlsread(filename_y3, sheet, xlRange);
            real_z3 = xlsread(filename_z3, sheet, xlRange);
            real_X1 = xlsread(filename_X1, sheet, xlRange);
            real_Y1 = xlsread(filename_Y1, sheet, xlRange);
            real_Z1 = xlsread(filename_Z1, sheet, xlRange);
            real_X2 = xlsread(filename_X2, sheet, xlRange);
            real_Y2 = xlsread(filename_Y2, sheet, xlRange);
            real_Z2 = xlsread(filename_Z2, sheet, xlRange);
            real_alpha = xlsread(filename_alpha, sheet, xlRange);

            for i = 1 : 5 : M
                for j = 1 : 10 : N
                    y0 = dz * abs(sin(atan(real_Z0(i, j) / real_X0(i, j))));
                    line([x0, real_X0(i, j)], [y0, real_Y0(i, j)], [z0, real_Z0(i, j)], 'color', CMP(256 - i * floor(256 / M), :)); hold on;
                    line([real_X0(i, j), real_X1(i, j)], [real_Y0(i, j), real_Y1(i, j)], [real_Z0(i, j), real_Z1(i, j)],'linewidth', 1, 'color', CMP(256 - i * floor(256 / M), :)); hold on;
                    line([real_X1(i, j), real_X2(i, j)], [real_Y1(i, j), real_Y2(i, j)], [real_Z1(i, j), real_Z2(i, j)],'linewidth', 1, 'color', CMP(256 - i * floor(256 / M), :)); hold on;
                    line([real_X2(i, j), real_x3(i, j)], [real_Y2(i, j), real_y3(i, j)], [real_Z2(i, j), real_Z3(i, j)],'linewidth', 1, 'color', CMP(256 - i * floor(256 / M), :)); hold on;
                    
                end
            end
        end
    end