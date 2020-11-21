unit Unit1;

interface

uses
 Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, TeEngine,ComObj, Series, ExtCtrls, TeeProcs, Chart,
  Menus;

type
  TForm1 = class(TForm)
    dlgSave1: TSaveDialog;
    dlgOpen1: TOpenDialog;
    mm1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    StringGrid: TStringGrid;
    lbl1: TLabel;
    lbl2: TLabel;
    lbl3: TLabel;
    lbl4: TLabel;
    lbl5: TLabel;
    lbl6: TLabel;
    edt1: TEdit;
    edt2: TEdit;
    edt3: TEdit;
    edt4: TEdit;
    edt5: TEdit;
    edt6: TEdit;
    btn1: TButton;
    cht1: TChart;
    lbl7: TLabel;
    edt7: TEdit;
    Series1: TFastLineSeries;
    procedure N4Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}
    var
  u: Real;
  g1:real;
  OTP: real;


procedure TForm1.N4Click(Sender: TObject);
begin
Close;
end;

procedure TForm1.N2Click(Sender: TObject);
  var
  Exdoc:Variant ;
  i,j:Integer;
begin
    Exdoc:= CreateOleObject ('Excel.Application');
    with dlgOpen1 do
    if Execute then
    Exdoc.application.WorkBooks.Add(FileName);
    with StringGrid do
    for i:= 1 to RowCount-1 do
    for j:= 1 to ColCount-1 do
    Cells[j,i]:=Exdoc.Cells[i,j];
    try
      Exdoc.Quit;
      except
        end;
    Exdoc:=Unassigned ;
begin with StringGrid do
  begin
    StringGrid.Cells[1,0]:= 'Здоровые';
    StringGrid.Cells[2,0]:= 'Больные';
    end;
end;
end;

procedure TForm1.btn1Click(Sender: TObject);
var
  step,OTP,max,min,interval,q,a,b:Real;
  P1:Extended;
  i,j,h,m,v,colichestvo,q1,u,z,len,summa1,summa2,j1,j2,zd,b1,g:Integer;

  PD: array of Integer;
  ND: array of Integer;
  PH: array of Integer;
  NH: array of Integer;
  OTP1: array of Real;
  Se: array of Real;
  Sp: array of Real;
begin
  //определяем max здоровых и min больных

  //max здоровых
  with StringGrid.Cols[1] do
   begin
     max:= StrToFloat(StringGrid.Cells[1,1]);
     for j:=1 to StringGrid.RowCount-1 do
       begin
         if StringGrid.Cells[1,j]<> '' then
           begin
             if max<StrToFloat(StringGrid.Cells[1,j]) then max:=StrToFloat(StringGrid.Cells[1,j]);
           end;
       end;
   end;
    edt3.Text:=FloatToStr(max);


    //min  больных
    with StringGrid.Cols[2] do
    begin
      min:=StrToFloat(StringGrid.Cells[2,1]);
      for h:=1 to StringGrid.RowCount-1 do
        begin
          if StringGrid.Cells[2,h]<> '' then
            begin
              if min>StrToFloat(StringGrid.Cells[2,h]) then min:= StrToFloat(StringGrid.Cells[2,h]);

            end;
        end;
    end;
    edt4.Text:=FloatToStr(min);
    interval:=max-min;
    edt6.Text:=FloatToStr(interval);


    // вычисления шага сдвига

    if max<min
    then
    begin OTP:=(max+min)/2;
     edt2.text:= FloatToStr(OTP);
    end;

    if (max>=min) then
    begin
    colichestvo:=0;  //обнулениеколичества
    for v:= 1 to StringGrid.RowCount-1 do
    for m:= 1 to StringGrid.ColCount-1 do
      begin
        if StringGrid.Cells[m,v]<> '' then
           begin
             q:= StrToFloat(StringGrid.Cells[m,v]);
             if (q>=min) and (q<max)
             then Inc(colichestvo); //счетчик количества элементов в области перекрытия
              edt5.Text:=IntToStr(colichestvo);
           end;

      end;

      step:=interval/colichestvo;
      edt1.Text:=FloatToStr(step);

         //Определение Оптимальной точки разделения ОТР
      SetLength(OTP1,colichestvo);
      OTP:=min;
      len:=0;
      j:=0;

       while (OTP<(max-0.01)) do
         begin
           OTP1[j]:=OTP;
           OTP:=OTP + step;
           len:= len+1;
           Inc(j);
         end;
        SetLength(PD,len-1);
        SetLength(ND,len-1);
        SetLength(PH,len-1);
        SetLength(NH,len-1);
        for j:=0 to len-2 do
          begin
            with StringGrid.Cols[1] do
              begin
                for v:=1 to StringGrid.RowCount-1 do
                  begin
                      if StringGrid.Cells[1,v]<> '' then
                        begin
                           a:= StrToFloat(StringGrid.Cells[1,v]);
                           if (a>=OTP1[j]) then PH[j]:=PH[j]+1;
                        end;
                  end;
              end;

              with StringGrid.Cols[2] do
                begin
                  for v:= 1 to StringGrid.RowCount-1 do
                    begin
                       if StringGrid.Cells[2,v]<> '' then
                         begin
                             b:= StrToFloat(StringGrid.Cells[2,v]);
                             if (b<OTP1[j]) then ND[j]:=ND[j]+1;
                         end;
                    end;
                end;
          end;


          summa1:= PH[0]+ND[0];
            for j:= 0 to len-1 do
             begin
                summa2:=PH[j]+ND[j];
                if (summa1>=summa2) then
                  begin
                     summa1:=summa2; //минимальная сумма ошибок
                     OTP:=OTP1[j];
                  end;
             end;
           edt2.text:=FloatToStr(OTP);

          zd:=0; b1:=0;
           for g:=1 to StringGrid.RowCount-1 do
              begin
                g1:=StrToFloat(StringGrid.Cells[1,g]);
                zd:=Zd+1;
                u:=StrToInt(StringGrid.Cells[2,g]);
                b1:=b1+1;
              end;



          //Вычисление чувствительности SE  и cпецифичности Sp, построение графика
       Series1.Clear;
       SetLength(Se,len-1);
       SetLength(Sp,len-1);
           for j2:=0 to len-2 do
             begin
                NH[j2]:=zd-PH[j2];
                PD[j2]:=b1-ND[j2];
                Se[j2]:=PD[j2]/(PD[j2]+ND[j2]);
                Sp[j2]:=1-(NH[j2]/(PH[j2]+NH[j2]));
             end;
           for j:=0 to len-1 do
              begin
                Series1.AddXY(Sp[j],Se[j], '' ,clgreen);
              end;
           //Вычисление площади
       P1:=0;
         for j1:=0 to len-1 do
           begin
             P1:=P1+(((1-Se[j1+1])+(1-Se[j1]))/2)*(Sp[j1]-Sp[j1+1]);
           end;
       P1:=1-P1;
       edt7.Text:=FloatToStrF(P1,ffFixed,10,3);
    end;


end;
procedure TForm1.FormCreate(Sender: TObject);
begin

end;

end.
