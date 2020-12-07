clear all; close all;
warning off; clc;

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
    PSSE_UUT_VOLTAGE        = Results_Psse.Out(:,4);
    PSSE_POC_VOLTAGE        = Results_Psse.Out(:,5);
    PSSE_UUT_PELEC          = Results_Psse.Out(:,6);
    PSSE_UUT_QELEC          = Results_Psse.Out(:,7);
    PSSE_P_POC              = Results_Psse.Out(:,8);
    PSSE_Q_POC              = Results_Psse.Out(:,9);
    PSSE_UUT_ID_CMD         = Results_Psse.Out(:,10);
    PSSE_UUT_IQ_CMD         = Results_Psse.Out(:,11);
    PSSE_PPC_P_CMD          = Results_Psse.Out(:,12);
    PSSE_PPC_Q_CMD          = Results_Psse.Out(:,13);
    
    % Plot Figures %
    h = figure;
    subplot(4,2,1)
    plot(PSSE_TIME,PSSE_UUT_VOLTAGE, 'r', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('VOLTAGE (pu)');
    title('UUT_VOLTAGE','Interpreter','none');
    
    subplot(4,2,2)
    PSSE_POC_VREF_0 = PSSE_POC_VOLTAGE(1,1);
    PSSE_POC_VREF_01 = 0.95*PSSE_POC_VOLTAGE(1,1);
    PSSE_POC_VREF_10 = 1.00*PSSE_POC_VOLTAGE(1,1);
    PSSE_POC_VREF_20 = 1.05*PSSE_POC_VOLTAGE(1,1);
    PSSE_POC_VREF_30 = 1.00*PSSE_POC_VOLTAGE(1,1);
    index_01 = find(abs(PSSE_TIME-1) <= 0.002);
    index_10 = find(abs(PSSE_TIME-10) <= 0.002);
    index_20 = find(abs(PSSE_TIME-20) <= 0.002);
    index_30 = find(abs(PSSE_TIME-30) <= 0.002);
    index_40 = find(abs(PSSE_TIME-40) <= 0.002);
    PSSE_POC_VREF = [ones(index_01(1,1),1)*PSSE_POC_VREF_0;...
        ones(index_10(1,1)-index_01(1,1),1)*PSSE_POC_VREF_01;...
        ones(index_20(1,1)-index_10(1,1),1)*PSSE_POC_VREF_10;...
        ones(index_30(1,1)-index_20(1,1),1)*PSSE_POC_VREF_20;...
        ones(index_40(end,1)-index_30(1,1),1)*PSSE_POC_VREF_30];
    plot(PSSE_TIME,PSSE_POC_VOLTAGE, 'r', 'LineWidth', 1.5); axis auto; hold on;
    plot(PSSE_TIME,PSSE_POC_VREF, ':k', 'LineWidth', 1.5); hold off;
    set(gca,'xlim',[0 tRun]); set(gca,'ylim',[0.90 1.10]); grid on;
    ylabel('VOLTAGE (pu)');
    title('POC_VOLTAGE','Interpreter','none');
    legend('POC VOLTAGE','POC VREF','Location','Best')
    
    subplot(4,2,3)
    plot(PSSE_TIME,PSSE_UUT_PELEC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('ACTIVE POWER (pu)');
    title('UUT_PELEC','Interpreter','none');
    
    subplot(4,2,4)
    plot(PSSE_TIME,PSSE_P_POC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('ACTIVE POWER (MW)');
    title('POC_P','Interpreter','none');
    
    subplot(4,2,5)
    plot(PSSE_TIME,PSSE_UUT_QELEC, 'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('REACTIVE POWER (pu)');
    title('UUT_QELEC','Interpreter','none');
    
    subplot(4,2,6)
    plot(PSSE_TIME,PSSE_Q_POC,'b', 'LineWidth', 1.5); axis tight;
    set(gca,'xlim',[0 tRun]); grid on;
    ylabel('REACTIVE POWER (MVAr)');
    title('POC_Q','Interpreter','none');
    
    subplot(4,2,7)
    plot(PSSE_TIME,PSSE_PPC_P_CMD,'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    xlabel('Time (s)'); ylabel('ID (pu)');
    title('UUT_ID_CMD','Interpreter','none');
    
    subplot(4,2,8)
    plot(PSSE_TIME,PSSE_PPC_Q_CMD,'b', 'LineWidth', 1.5); axis auto;
    set(gca,'xlim',[0 tRun]); grid on;
    xlabel('Time (s)'); ylabel('IQ (pu)');
    title('UUT_IQ_CMD','Interpreter','none');
    
    set(gcf,'Position',[200,200,1200,760], 'color','w')
    temp = strrep(User_File(1:end-4),'_','\_');
    suptitle(temp);
    print(gcf,'-dmeta', '-r360', sprintf('%s.emf', User_File(1:end-4)));
    close all;
end