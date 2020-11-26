clear all; close all;
warning off; clc;
Excel_File = '06. PoC Qref Step Test.xlsx';
List = dir('*.out');
delete('*.emf');
[row,col] = size(List);
tRun = 30;

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
    PSSE_PPC_P_CMD          = Results_Psse.Out(:,12);
    PSSE_PPC_Q_CMD          = Results_Psse.Out(:,13);
    
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
    temp = mean(PSSE_INV_PELEC); temp_max = temp*1.2; temp_min = temp*0.8;
    set(gca,'xlim',[0 tRun]); grid on;
    set(gca,'ylim',[temp_min temp_max]);
    ylabel('ACTIVE POWER (MW)');
    title('INV_PELEC','Interpreter','none');
    
    subplot(4,2,4)
    plot(PSSE_TIME,PSSE_P_POC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    set(gca,'ylim',[temp_min temp_max]);
    ylabel('ACTIVE POWER (MW)');
    title('POC_P','Interpreter','none');
    
    subplot(4,2,5)
    plot(PSSE_TIME,PSSE_INV_QELEC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('REACTIVE POWER (MVAr)');
    title('INV_QELEC','Interpreter','none');

    subplot(4,2,6)    
%     PSSE_POC_QREF_0 = PSSE_Q_POC(1,1);
%     PSSE_POC_QREF_01 = PSSE_Q_POC(1,1) - 0.3*100;
%     PSSE_POC_QREF_10 = PSSE_Q_POC(1,1) + 0.3*100;
%     PSSE_POC_QREF_20 = PSSE_Q_POC(1,1);
%     index_01 = find(abs(PSSE_TIME-1) <= 0.002);
%     index_10 = find(abs(PSSE_TIME-10) <= 0.002);
%     index_20 = find(abs(PSSE_TIME-20) <= 0.002);
%     index_30 = find(abs(PSSE_TIME-30) <= 0.002);
%     PSSE_POC_VREF = [ones(index_01(1,1),1)*PSSE_POC_QREF_0;...
%         ones(index_10(1,1)-index_01(1,1),1)*PSSE_POC_QREF_01;...
%         ones(index_20(1,1)-index_10(1,1),1)*PSSE_POC_QREF_10;...
%         ones(index_30(1,1)-index_20(1,1),1)*PSSE_POC_QREF_20];
    plot(PSSE_TIME,PSSE_Q_POC,'b', 'LineWidth', 1.5); axis auto; hold on;
%     plot(PSSE_TIME,PSSE_POC_VREF, ':r', 'LineWidth', 1.5); hold off;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('REACTIVE POWER (MVAr)');
    title('POC_Q','Interpreter','none');
    legend('POC REACTIVE POWER','POC QREF','Location','Best')
    
    subplot(4,2,7)
    plot(PSSE_TIME,PSSE_PPC_P_CMD,'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    xlabel('Time (s)'); ylabel('ID (pu)');
    title('INV_ID_CMD','Interpreter','none');
    
    subplot(4,2,8)
    plot(PSSE_TIME,PSSE_PPC_Q_CMD,'b', 'LineWidth', 1.5); axis auto;
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