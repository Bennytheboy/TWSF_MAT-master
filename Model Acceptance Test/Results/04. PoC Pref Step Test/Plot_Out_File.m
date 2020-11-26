clear all; close all;
warning off; clc;
Excel_File = '04. PoC Pref Step Test.xlsx';
List = dir('*.out');
delete('*.emf');
[row,col] = size(List);
tRun = 40;

for i = 1:row
    % Load PSS/E Simulation Results %
    User_File = List(i,1).name;
    Results_Psse = Read_Out_File(User_File);
    PSSE_TIME               = Results_Psse.Out(:,1);
    PSSE_SYS_FREQ_DEVIATION = Results_Psse.Out(:,2);
    PSSE_IB_VOLTAGE         = Results_Psse.Out(:,3);
    PSSE_INV_VOLTAGE        = Results_Psse.Out(:,4);
    PSSE_POC_VOLTAGE        = Results_Psse.Out(:,5);
    PSSE_INV_PELEC          = Results_Psse.Out(:,6)*100;
    PSSE_INV_QELEC          = Results_Psse.Out(:,7)*100;
    PSSE_P_POC              = Results_Psse.Out(:,8);
    PSSE_Q_POC              = Results_Psse.Out(:,9);
    PSSE_INV_ID_CMD         = Results_Psse.Out(:,10);
    PSSE_INV_IQ_CMD         = Results_Psse.Out(:,11);
    
    % Plot Figures %
    h = figure;
    subplot(4,2,1)
    plot(PSSE_TIME,PSSE_INV_VOLTAGE, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('VOLTAGE (pu)');
    title('INV_VOLTAGE','Interpreter','none');
    
    subplot(4,2,2)
    plot(PSSE_TIME,PSSE_POC_VOLTAGE, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('VOLTAGE (pu)');
    title('POC_VOLTAGE','Interpreter','none');
    
    subplot(4,2,3)
    plot(PSSE_TIME,PSSE_INV_PELEC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('ACTIVE POWER (MW)');
    title('INV_PELEC','Interpreter','none');
    
    subplot(4,2,4)
    PSSE_POC_PREF_0 = PSSE_P_POC(1,1);
    PSSE_POC_PREF_05 = 0.8*PSSE_P_POC(1,1);
    PSSE_POC_PREF_10 = 0.6*PSSE_P_POC(1,1);
    PSSE_POC_PREF_15 = 0.4*PSSE_P_POC(1,1);
    PSSE_POC_PREF_20 = 0.2*PSSE_P_POC(1,1);
    PSSE_POC_PREF_25 = PSSE_P_POC(1,1);
    index_5 = find(abs(PSSE_TIME-5) <= 0.002);
    index_10 = find(abs(PSSE_TIME-10) <= 0.002);
    index_15 = find(abs(PSSE_TIME-15) <= 0.002);
    index_20 = find(abs(PSSE_TIME-20) <= 0.002);
    index_25 = find(abs(PSSE_TIME-25) <= 0.002);
    index_50 = size(PSSE_TIME-25,1);
    PSSE_POC_PREF = [ones(index_5(1,1),1)*PSSE_POC_PREF_0;...
        ones(index_10(1,1)-index_5(1,1),1)*PSSE_POC_PREF_05;...
        ones(index_15(1,1)-index_10(1,1),1)*PSSE_POC_PREF_10;...
        ones(index_20(1,1)-index_15(1,1),1)*PSSE_POC_PREF_15;...
        ones(index_25(1,1)-index_20(1,1),1)*PSSE_POC_PREF_20;...
        ones(index_50(end,1)-index_25(1,1),1)*PSSE_POC_PREF_25];
    plot(PSSE_TIME,PSSE_P_POC, 'b', 'LineWidth', 1.5); axis auto; hold on;
    plot(PSSE_TIME,PSSE_POC_PREF, ':r', 'LineWidth', 1.5); hold off;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('ACTIVE POWER (MW)');
    title('POC_P','Interpreter','none');
    legend('POC ACTIVE POWER','POC PREF','Location','Best')
    
    subplot(4,2,5)
    plot(PSSE_TIME,PSSE_INV_QELEC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('REACTIVE POWER (MVAr)');
    title('INV_QELEC','Interpreter','none');
    
    subplot(4,2,6)
    plot(PSSE_TIME,PSSE_Q_POC,'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('REACTIVE POWER (MVAr)');
    title('POC_Q','Interpreter','none');
    
    subplot(4,2,7)
    plot(PSSE_TIME,PSSE_INV_ID_CMD,'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    xlabel('Time (s)'); ylabel('ID (pu)');
    title('INV_ID_CMD','Interpreter','none');
    
    subplot(4,2,8)
    plot(PSSE_TIME,PSSE_INV_IQ_CMD,'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    xlabel('Time (s)'); ylabel('IQ (pu)');
    title('INV_IQ_CMD','Interpreter','none');
    
    set(gcf,'Position',[200,200,1200,760], 'color','w')
    temp = strrep(User_File(1:end-4),'_','\_');
    suptitle(temp);
    print(gcf,'-dmeta', '-r360', sprintf('%s.emf', User_File(1:end-4)));
%     xlswritefig(h,Excel_File,'Sheet1',sprintf('%s%d','B',(i-1)*45+1));
    close all;
end