clear all; close all; clc;
format longE
formatSpec = '%.15f';
M = 150; N = 50;
x0 = 0; y0 = 0; z0 = 0;
tic
%% Ҫȷ���Ĳ�����
L1_start = 2; L1_interval = 0.1; L1_end = 9; %��һ��͸����λ�ò���
L2_start = 5; L2_interval = 0.1; L2_end = 10;%�ڶ���͸����λ�ò���
a1 = 20;an = 25;d = 1;%�������߰뾶
%% ���
for w = 25  %: d : an
    for L1 = L1_start %: L1_interval : L1_end
        for L2 = L2_start %: L2_interval : L2_end
            sheet = round((L1 - L1_start) / L1_interval) + 1; %�ⲿѭ��????������һ��sheet
            pdx = round((L2 - L2_start) / L2_interval) + 1; %��¼�ڲ�ѭ������
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
            Z0 = xlsread(filename_Z0, sheet, xlRange); %����������񻮷�����
            X3 = xlsread(filename_X3, sheet, xlRange);
            Y3 = xlsread(filename_Y3, sheet, xlRange);
            Z3 = xlsread(filename_Z3, sheet, xlRange); %�����ߵ����񻮷���??
            x1 = xlsread(filename_xx1, sheet, xlRange);
            y1 = xlsread(filename_yy1, sheet, xlRange);
            z1 = xlsread(filename_zz1, sheet, xlRange);
            x2 = xlsread(filename_xx2, sheet, xlRange);
            y2 = xlsread(filename_yy2, sheet, xlRange);
            z2 = xlsread(filename_zz2, sheet, xlRange);
            figure
            axis equal
            cdata = cat(3, zeros(size(X0)), zeros(size(X0)), zeros(size(X0))); %������ʹ�ұߵĳ������ɺ�ɫ
            mesh(X3, Y3, Z3, cdata); hold on %����??
            plot3(X3, Y3, Z3, 'o', 'markersize', 4, 'markerfacecolor', 'k', 'markeredgecolor', 'k'); hold on
            plot3(X0, Y0, Z0, 'o', 'markersize', 4, 'markerfacecolor', 'k', 'markeredgecolor', 'k'); hold on
            xlabel('x'); ylabel('y'); zlabel('z');
        end
    end
end