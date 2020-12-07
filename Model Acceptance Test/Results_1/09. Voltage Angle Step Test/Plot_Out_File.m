clear all; close all;
warning off; clc;
Excel_File = '08. Voltage Angle Step Test.xlsx';
List = dir('*.out');
delete('*.emf');
[row,col] = size(List);
tRun = 25;

for i = 1:row
    % Load PSS/E Simulation Results %
    User_File = List(i,1).name;
    Results_Psse = Read_Out_File(User_File);
    PSSE_TIME               = Results_Psse.Out(:,1);
    PSSE_SYS_FREQ_DEVIATION = Results_Psse.Out(:,2);
    PSSE_IB_VOLTAGE         = Results_Psse.Out(:,3);
    PSSE_INV_VOLTAGE        = Results_Psse.Out(:,4);
    PSSE_INV_VOLTAGE_ANGLE  = Results_Psse.Out(:,5);
    PSSE_POC_VOLTAGE        = Results_Psse.Out(:,6);
    PSSE_POC_VOLTAGE_ANGLE  = Results_Psse.Out(:,7);
    PSSE_INV_PELEC          = Results_Psse.Out(:,8)*100;
    PSSE_INV_QELEC          = Results_Psse.Out(:,9)*100;
    PSSE_P_POC              = Results_Psse.Out(:,10);
    PSSE_Q_POC              = Results_Psse.Out(:,11);
    PSSE_INV_ID_CMD         = Results_Psse.Out(:,12);
    PSSE_INV_IQ_CMD         = Results_Psse.Out(:,13);
    
    % Plot Figures %
    h = figure;
    subplot(4,2,1)
    plot(PSSE_TIME,PSSE_INV_VOLTAGE,'b', 'LineWidth', 1.5); hold on;
    plot(PSSE_TIME,PSSE_POC_VOLTAGE,'r', 'LineWidth', 1.5); hold off; axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('VOLTAGE (pu)');
    title('VOLTAGE MAGNITUDE','Interpreter','none');
    legend('INV_VOLTAGE','POC_VOLTAGE','Location','Best','Interpreter','none')
    
    subplot(4,2,2)
    plot(PSSE_TIME,PSSE_INV_VOLTAGE_ANGLE,'b', 'LineWidth', 1.5); hold on;
    plot(PSSE_TIME,PSSE_POC_VOLTAGE_ANGLE,'r', 'LineWidth', 1.5); hold off; axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('ANGLE (o)');
    title('VOLTAGE ANGLE','Interpreter','none');
    legend('INV_VOLTAGE_ANGLE','POC_VOLTAGE_ANGLE','Location','Best','Interpreter','none')
    
    subplot(4,2,3)
    plot(PSSE_TIME,PSSE_INV_PELEC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('ACTIVE POWER (pu)');
    title('INV_PELEC','Interpreter','none');
    
    subplot(4,2,4)
    plot(PSSE_TIME,PSSE_P_POC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('ACTIVE POWER (MW)');
    title('POC_P','Interpreter','none');
    
    subplot(4,2,5)
    plot(PSSE_TIME,PSSE_INV_QELEC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('REACTIVE POWER (pu)');
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