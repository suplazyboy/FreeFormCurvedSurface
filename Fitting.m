clear all; close all; clc;
format longE
formatSpec = '%.15f';
M = 20; N = 20;dz = 0.3;
n1 = 1; n2 = 1.51; n3 = 1;
x0 = 0; y0 = 0; z0 = 0;ze = 0;
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
            dictionaryname = ['w=', num2str(w), ' L1=', num2str(L1), ' L2=', num2str(L2)];
            filename_X0 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\X0.xlsx'];
            filename_Y0 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\Y0.xlsx'];
            filename_Z0 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\Z0.xlsx'];
            filename_X3 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\X3.xlsx'];
            filename_Y3 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\Y3.xlsx'];
            filename_Z3 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\Z3.xlsx'];
            filename_xx1 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\xx1.xlsx'];
            filename_yy1 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\yy1.xlsx'];
            filename_zz1 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\zz1.xlsx'];
            filename_xx2 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\xx2.xlsx'];
            filename_yy2 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\yy2.xlsx'];
            filename_zz2 = ['F:\FreeFormAssets\CurvedSurface\',dictionaryname '\zz2.xlsx'];
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

            for j = 1 : M
                for i = 1 : N
                    if (X0(j, i) == 0)
                        y0 = 0;
                    else
                        y0 = dz * abs(sin(atan(Z0(j, i) / X0(j, i))));
                    end
                    line([x0, X0(j, i)], [y0, Y0(j, i)], [z0, Z0(j, i)], 'color', 'r'); hold on;
                    I1{j, i}(1) = X0(j, i) - x0;
                    I1{j, i}(2) = Y0(j, i) - y0;
                    I1{j, i}(3) = Z0(j, i) - z0;
                    I1{j, i} = normr(I1{j, i});

                    %yi表示的是第一个镜上的�?
                    options_1=optimset('tolfun',1e-6,'Display','off');
                    fun_1=@(yi)[(yi(2)-y0).*I1{j,i}(1)./I1{j,i}(2)+x0-yi(1);...
                        (yi(2)-y0).*I1{j,i}(3)./I1{j,i}(2)+z0-yi(3);...
                        p(1) + p(2).*yi(1) + p(3).*yi(3) + p(4).*yi(1).^2 + p(5).*yi(1).*yi(3) + p(6).*yi(3).^2 + p(7).*yi(1).^3 + p(8).*yi(1).^2.*yi(3)+...
                        p(9).*yi(1).*yi(3).^2 + p(10).*yi(3).^3 + p(11).*yi(1).^4 + p(12).*yi(1).^3.*yi(3) + p(13).*yi(1).^2.*yi(3).^2+...
                        p(14).*yi(1).*yi(3).^3 + p(15).*yi(3).^4 + p(16).*yi(1).^5 + p(17).*yi(1).^4.*yi(3) + p(18).*yi(1).^3.*yi(3).^2+...
                        p(19).*yi(1).^2.*yi(3).^3 + p(20).*yi(1).*yi(3).^4 + p(21).*yi(3).^5-yi(2)];
                    yi = fsolve(fun_1, [x1(j, i); y1(j, i); z1(j, i)], options_1);
                    X1(j, i) = yi(1); Y1(j, i) = yi(2); Z1(j, i) = yi(3);
                    line([X0(j, i), X1(j, i)], [Y0(j, i), Y1(j, i)], [Z0(j, i), Z1(j, i)], 'linewidth', 1, 'color', 'r'); hold on
                    %N1表示的是第一个镜上的法向矢量
                    syms x y z
                    F1 = p(1) + p(2) .* x + p(3) .* z + p(4) .* x .^ 2 + p(5) .* x .* z + p(6) .* z .^ 2 + p(7) .* x .^ 3 + p(8) .* x .^ 2 .* z + ...
                        p(9) .* x .* z .^ 2 + p(10) .* z .^ 3 + p(11) .* x .^ 4 + p(12) .* x .^ 3 .* z + p(13) .* x .^ 2 .* z .^ 2 + ...
                        p(14) .* x .* z .^ 3 + p(15) .* z .^ 4 + p(16) .* x .^ 5 + p(17) .* x .^ 4 .* z + p(18) .* x .^ 3 .* z .^ 2 + ...
                        p(19) .* x .^ 2 .* z .^ 3 + p(20) .* x .* z .^ 4 + p(21) .* z .^ 5 - y;
                    Fx1 = diff(F1, x);
                    Fz1 = diff(F1, z);
                    x = yi(1); z = yi(3);
                    N1{j, i}(1) = double(eval(Fx1));
                    N1{j, i}(2) = -1;
                    N1{j, i}(3) = double(eval(Fz1));
                    N1{j, i} = normr(N1{j, i});
                    %line([X1(j, i), X1(j, i) + N1{j, i}(1)],[Y1(j, i), Y1(j, i) + N1{j, i}(2)], [Z1(j, i), Z1(j, i) + N1{j, i}(3)]);
                    axis equal;
                    %er表示的是第一个镜的出射光线矢�?
                    options_2=optimset('tolfun',1e-6,'Display','off');
                    fun_2=@(er)(n2 * er - n1 * I1{j, i} - (n2 * dot(N1{j, i}, er) - n1 * dot(N1{j, i}, I1{j, i})) .* N1{j, i});
                    er=fsolve(fun_2,[x2(j,i)-x1(j,i) y2(j,i)-y1(j,i) z2(j,i)-z1(j,i)],options_2);
                    O1{j, i}(1) = er(1); O1{j, i}(2) = er(2); O1{j, i}(3) = er(3); %次镜的出射光线矢�?
                    I2{j, i} = normr(O1{j, i});
                    %line([X1(j, i), X1(j, i) + I2{j, i}(1) .* 3], [Y1(j, i), Y1(j, i) + I2{j, i}(2) .* 3], [Z1(j, i), Z1(j, i) + I2{j, i}(3) .* 3], 'linewidth', 1, 'color', 'g'); hold on
                end 
            end

            for j = 1 : M
                for i = 1 : N
                    %san表示的是第二个镜上的�?
                    options_3 = optimset('TolFun', 1e-6, 'Display', 'off');
                    fun_3 = @(san)[(san(2) - Y1(j, i)) .* I2{j, i}(3) ./ I2{j, i}(2) + Z1(j, i) - san(3); ...
                        (san(2) - Y1(j, i)) .* I2{j, i}(1) ./ I2{j, i}(2) + X1(j, i) - san(1); ...
                        q(1) + q(2) .* san(1) + q(3) .* san(3) + q(4) .* san(1) .^ 2 + q(5) .* san(1) .* san(3) + q(6) .* san(3) .^ 2 + q(7) .* san(1) .^ 3 + q(8) .* san(1) .^ 2 .* san(3) + ...
                        q(9) .* san(1) .* san(3) .^ 2 + q(10) .* san(3) .^ 3 + q(11) .* san(1) .^ 4 + q(12) .* san(1) .^ 3 .* san(3) + q(13) .* san(1) .^ 2 .* san(3) .^ 2 + ...
                        q(14) .* san(1) .* san(3) .^ 3 + q(15) .* san(3) .^ 4 + q(16) .* san(1) .^ 5 + q(17) .* san(1) .^ 4 .* san(3) + q(18) .* san(1) .^ 3 .* san(3) .^ 2 + ...
                        q(19) .* san(1) .^ 2 .* san(3) .^ 3 + q(20) .* san(1) .* san(3) .^ 4 + q(21) .* san(3) .^ 5 - san(2)];
                    san = fsolve(fun_3, [x2(j, i); y2(j, i); z2(j, i)], options_3);
                    X2(j, i) = san(1); Y2(j, i) = san(2); Z2(j, i) = san(3);
                    line([X1(j, i), X2(j, i)], [Y1(j, i), Y2(j, i)], [Z1(j, i), Z2(j, i)], 'linewidth', 1, 'color', 'g'); hold on
                    %N2表示的是第二个镜上的法向矢量
                    syms x y z
                    F1 = q(1) + q(2) .* x + q(3) .* z + q(4) .* x .^ 2 + q(5) .* x .* z + q(6) .* z .^ 2 + q(7) .* x .^ 3 + q(8) .* x .^ 2 .* z + ...
                        q(9) .* x .* z .^ 2 + q(10) .* z .^ 3 + q(11) .* x .^ 4 + q(12) .* x .^ 3 .* z + q(13) .* x .^ 2 .* z .^ 2 + ...
                        q(14) .* x .* z .^ 3 + q(15) .* z .^ 4 + q(16) .* x .^ 5 + q(17) .* x .^ 4 .* z + q(18) .* x .^ 3 .* z .^ 2 + ...
                        q(19) .* x .^ 2 .* z .^ 3 + q(20) .* x .* z .^ 4 + q(21) .* z .^ 5 - y;
                    Fx1 = diff(F1, x);
                    Fz1 = diff(F1, z);
                    x = san(1); z = san(3);
                    N2{j, i}(1) = double(eval(Fx1));
                    N2{j, i}(3) = double(eval(Fz1));
                    N2{j, i}(2) = -1;
                    N2{j, i} = normr(N2{j, i});

                    %si表示的是主镜的出射光线矢�?
                    options_4=optimset('tolfun',1e-6,'Display','off');
                    
                    fun_4 = @(si)(n3 * si - n2 * I2{j, i} - (n3 * dot(N2{j, i}, si) - n3 * dot(N2{j, i}, I2{j, i})) .* N2{j, i});
                    si = fsolve(fun_4,[X3(j,i)-x2(j,i) Y3(j,i)-y2(j,i) Z3(j,i)-z2(j,i)],options_2);
                    O2{j,i}(1)=si(1);O2{j,i}(2)=si(2);O2{j,i}(3)=si(3);%主镜的出射光线矢�?
                    I3{j, i} = normr(O2{j,i});
                    %计算发散�?
                    A3 = [0, 1, 0];
                    alpha_mrad(j, i) = 1000 .* acos(dot(O2{j, i}, A3) / norm(O2{j, i}, 2) / norm(A3, 2));

                    %计算出射面上的点
                    options_5 = optimset('tolfun', 1e-6, 'Display', 'off');
                    fun_5 = @(wu)[
                        (Y3(j, i) - Y2(j, i)) .* O2{j, i}(1) ./ O2{j, i}(2) + X2(j, i) - wu(1); ...
                        (Y3(j, i) - Y2(j, i)) .* O2{j, i}(3) ./ O2{j, i}(2) + Z2(j, i) - wu(2)
                    ];
                    wu = fsolve(fun_5, [X3(j, i) Z3(j, i)], options_3);
                    x3(j, i) = wu(1); y3(j, i) = Y3(j, i); z3(j, i) = wu(2);
                    line([X2(j, i), x3(j, i)], [Y2(j, i), y3(j, i)], [Z2(j, i), z3(j, i)], 'linewidth', 1, 'color', 'b'); hold on
                end
            end
            dictionaryname = ['w=', num2str(w), ' L1=', num2str(L1), ' L2=', num2str(L2)];
            pathToFolder = fullfile('F:\FreeFormAssets\RealPoint\', dictionaryname); % ��ȡ��ǰ����Ŀ¼�����ļ���������ϳ�����·��?
            if ~exist(pathToFolder, 'dir') % �жϸ�·�����Ƿ��Ѵ���ͬ���ļ���
                mkdir(pathToFolder); % ����mkdir()���������ļ���
            end
            filename_X0 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_X0.xlsx'];
            filename_Y0 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_Y0.xlsx'];
            filename_Z0 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_Z0.xlsx'];
            filename_x3 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_x3.xlsx'];
            filename_y3 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_y3.xlsx'];
            filename_z3 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_z3.xlsx'];
            filename_X1 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_X1.xlsx'];
            filename_Y1 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_Y1.xlsx'];
            filename_Z1 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_Z1.xlsx'];
            filename_X2 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_X2.xlsx'];
            filename_Y2 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_Y2.xlsx'];
            filename_Z2 = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\real_Z2.xlsx'];
            filename_alpha = ['F:\FreeFormAssets\RealPoint\', dictionaryname, '\alpha.xlsx'];
            xlswrite(filename_X0, X0, sheet, xlRange);
            xlswrite(filename_Y0, Y0, sheet, xlRange);
            xlswrite(filename_Z0, Z0, sheet, xlRange);
            xlswrite(filename_x3, x3, sheet, xlRange);
            xlswrite(filename_y3, y3, sheet, xlRange);
            xlswrite(filename_z3, z3, sheet, xlRange);
            xlswrite(filename_X1, X1, sheet, xlRange);
            xlswrite(filename_Y1, Y1, sheet, xlRange);
            xlswrite(filename_Z1, Z1, sheet, xlRange);
            xlswrite(filename_X2, X2, sheet, xlRange);
            xlswrite(filename_Y2, Y2, sheet, xlRange);
            xlswrite(filename_Z2, Z2, sheet, xlRange);
            xlswrite(filename_alpha, alpha_mrad, sheet, xlRange);
        end
    end
end
toc