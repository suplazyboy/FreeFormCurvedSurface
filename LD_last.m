clear all; close all; clc;
%���񻮷�
L0 = 1; dz = 0.3; L = 10;
faix = pi / 6; faiy = pi / 18;
N = 150; %�����������?�ԽǶ�t0���зָ�
M = 50; %�����������?�Գ���r0���зָ�
lambda = 1.55 * 10 ^ (-3);
zy0 = 0; zx0 = 0; z1 = L0; R0 = 0.5;
wx0 = lambda * 2 / pi ./ faix; wy0 = lambda * 2 / pi ./ faiy;
ZRx = pi .* wx0 .^ 2 ./ lambda; wx = wx0 .* sqrt(1 + ((z1 - zx0) / ZRx) .^ 2);
ZRy = pi .* wy0 .^ 2 ./ lambda; wy = wy0 .* sqrt(1 + ((z1 - zy0) / ZRy) .^ 2);
%% �����������������������������������������񡪡���������������������������������
syms r0 t0
a = wx; b = wy;
x = r0 * a * cos(t0); y = r0 * b * sin(t0);
f1 = 2 / (wx * wy * pi) * exp(-2 * x ^ 2 / wx ^ 2 - 2 * y ^ 2 / wy ^ 2);
sum_in = int(int(a * b * f1 * r0, r0, 0, 1), t0, 0, 2 * pi); sum_in = double(sum_in);
I0 = sum_in / ((N - 1) * (M - 1)); %������ÿһС�������?
R0(1) = 0; T0(1) = 0;%�������������� 
for i = 1 : M - 1
    syms r0 t0 R00
    x = r0 * a * cos(t0); y = r0 * b * sin(t0);
    [R00] = vpasolve(int(int(a * b * f1 * r0, t0, 0, 2 * pi), r0, R0(i), R00) - (N - 1) * I0);
    R0(i + 1) = double(abs(R00));
end
for i = 1 : N - 1
    syms r0 t0 T00
    x = r0 * a * cos(t0); y = r0 * b * sin(t0);
    [T00] = vpasolve(int(int(a * b * f1 * r0, r0, 0, 1), t0, T0(i), T00) - (M - 1) * I0);
    T0(i + 1) = double(abs(T00));
end
for i = 1 : M
    for j = 1 : N
        
        X0(i, j) = R0(i) * a * cos(T0(j));
        Z0(i, j) = R0(i) * b * sin(T0(j));
    end
end
Y0 = L0 * ones(size(X0));
cdata = cat(3, zeros(size(X0)), zeros(size(X0)), zeros(size(X0))); %��ɫ
filename_X0 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'X0_L0=' num2str(L0) '.xlsx'];
filename_Y0 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Y0_L0=' num2str(L0) '.xlsx'];
filename_Z0 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Z0_L0=' num2str(L0) '.xlsx'];
xlswrite(filename_X0, X0); %�����ļ�
xlswrite(filename_Y0, Y0); %�����ļ�
xlswrite(filename_Z0, Z0); %�����ļ�

%% ���������������������������ֳ��������񡪡���������������������������������
a1 = 20; d = 1; an = 25;
for w = 25 %a1 : d : an
    syms r0 t0
    x = r0 * cos(t0); y = r0 * sin(t0);
    f2 = 2 / (w ^ 2 * pi) * exp(-2 * (x ^ 2 + y ^ 2) / w ^ 2);
    R3(1) = 0; T3(1) = 0;
    sum_out = double(int(int(f2 * r0, r0, 0, w), t0, 0, 2 * pi));
    I3 = sum_out / ((N - 1) * (M - 1)); %������ÿһС�������?
    for i = 1 : M - 1
        syms r0 t0 R3_temp
        x = r0 * cos(t0); y = r0 * sin(t0);
        [R3_temp] = vpasolve(int(int(f2 * r0, t0, 0, 2 * pi), r0, R3(i), R3_temp) - (N - 1) * I3);
        R3(i + 1) = double(abs(R3_temp));
    end
    for i = 1 : N - 1
        syms r0 t0 T3_temp
        x = r0 * cos(t0); y = r0 * sin(t0);
        [T3_temp] = vpasolve(int(int(f2 * r0, r0, 0, w), t0, T3(i), T3_temp) - (M - 1) * I3);
        T3(i + 1) = double(abs(T3_temp));
    end
    for i = 1 : M
        for j = 1 : N
            X3(i, j) = R3(i) * cos(T3(j));
            Z3(i, j) = R3(i) * sin(T3(j));
        end
    end
    Y3 = L * ones(size(X3));
    sheet = (w - a1) / d + 1;
    filename_X3 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'X3_L3=' num2str(L) 'w=' num2str(a1) '-' num2str(d) '-' num2str(an) '.xlsx'];
    filename_Y3 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Y3_L3=' num2str(L) 'w=' num2str(a1) '-' num2str(d) '-' num2str(an) '.xlsx'];
    filename_Z3 = ['F:\FreeFormAssets\EnergyGrid\M' num2str(M) 'N' num2str(N) 'Z3_L3=' num2str(L) 'w=' num2str(a1) '-' num2str(d) '-' num2str(an) '.xlsx'];
    xlswrite(filename_X3, X3, sheet); %�����ļ�
    xlswrite(filename_Y3, Y3, sheet); %�����ļ�
    xlswrite(filename_Z3, Z3, sheet); %�����ļ�
end