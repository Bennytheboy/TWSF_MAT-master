function PSSE_Output = Read_Out_File(File_Name)
%% --- Load File --- %%
fid = fopen(File_Name);
setstr(fread(fid,4,'char')');
setstr(fread(fid,4,'char')');
setstr(fread(fid,4,'char')');
antal = fread(fid,1,'float32');
%% --- Read Data --- %%
fread(fid,1,'float32');
PSSE_Output.Channels = setstr(fread(fid,[32,antal],'char'))';
PSSE_Output.Title = setstr(fread(fid,[60,2] ,'char'))';
data = fread(fid,Inf,'float32');
tider = (size(data,1)-2)/(antal+2);
resultat = reshape(data(1:tider*(antal+2)),antal+2,tider)';
PSSE_Output.Out = resultat(:,2:size(resultat,2));
fclose(fid);