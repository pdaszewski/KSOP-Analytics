{
Unit z oknem głównym aplikacji.
Do poprawnego działania wymagana jest baza danych dostępna dla pracowników ZUOP, Urzęgu Gminy Różan i Krajowego Składowiska Odpadów Promienietwórczych w Różanie.
}

unit AOknoGl_frm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.ComCtrls, Vcl.ExtCtrls,
  Data.DB, Data.Win.ADODB, Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls,
  Vcl.Samples.Spin, Vcl.Imaging.jpeg, Vcl.Imaging.pngimage, Vcl.DBCtrls;

type
  TAOknoGl = class(TForm)
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    Plik1: TMenuItem;
    Zamknij1: TMenuItem;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    ADOConnection: TADOConnection;
    AutoRun: TTimer;
    ADOKarty: TADODataSet;
    DSKarty: TDataSource;
    grd_karty: TDBGrid;
    gbx_karta: TGroupBox;
    Label1: TLabel;
    btn_zapisz_karte: TButton;
    btn_nowa: TButton;
    edt_numer: TEdit;
    dtp_dostawa: TDateTimePicker;
    Label2: TLabel;
    Label3: TLabel;
    cmb_rodzaj: TComboBox;
    Label4: TLabel;
    Label5: TLabel;
    edt_objetosc: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    edt_waga: TEdit;
    Label8: TLabel;
    Label10: TLabel;
    cmb_kategoria: TComboBox;
    Label11: TLabel;
    cmb_podkategoria: TComboBox;
    gbx_izotopy: TGroupBox;
    img_logo: TImage;
    ADOQuery: TADOQuery;
    edt_szukaj: TEdit;
    DBGrid2: TDBGrid;
    DBNavigator1: TDBNavigator;
    ADOIzotopy: TADODataSet;
    DSIzotopy: TDataSource;
    ADOIzotopyid_poz_izotopu: TAutoIncField;
    ADOIzotopyid_karty: TIntegerField;
    ADOIzotopyizotop: TStringField;
    ADOIzotopyaktywnosc_poczatkowa: TFloatField;
    ADOIzotopyjednostka_akt_poczatkowej: TStringField;
    Label9: TLabel;
    edt_zrodel: TEdit;
    btn_czysc: TButton;
    Label12: TLabel;
    dtp_pomiar: TDateTimePicker;
    edt_symbol: TEdit;
    Pomoc1: TMenuItem;
    Oprogramie1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure Zamknij1Click(Sender: TObject);
    procedure AutoRunTimer(Sender: TObject);
    procedure btn_nowaClick(Sender: TObject);
    procedure Dodaj_nowa_karte_do_bazy;
    procedure btn_zapisz_karteClick(Sender: TObject);
    procedure edt_szukajKeyPress(Sender: TObject; var Key: Char);
    procedure Szukaj_karty;
    procedure edt_numerChange(Sender: TObject);
    procedure edt_objetoscKeyPress(Sender: TObject; var Key: Char);
    procedure Wczytaj_karte(id_karty : String);
    procedure grd_kartyDblClick(Sender: TObject);
    function Sprawdz_czy_karta_juz_istnieje(numer_karty : String): Boolean;
    procedure ADOIzotopyBeforePost(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure btn_czyscClick(Sender: TObject);
    procedure Aktualizuj_dane_karty;
    procedure ADOIzotopyNewRecord(DataSet: TDataSet);
    procedure Wypelnij_liste_izotopow;
    procedure DBGrid2CellClick(Column: TColumn);
    procedure DBGrid2ColEnter(Sender: TObject);
    procedure Oprogramie1Click(Sender: TObject);
    procedure cmb_rodzajChange(Sender: TObject);
    procedure edt_symbolChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
 wersja = '1.0.6';
var
  AOknoGl: TAOknoGl;
  wczytana_karta : string;

implementation

{$R *.dfm}

procedure TAOknoGl.ADOIzotopyBeforePost(DataSet: TDataSet);
begin
 ADOIzotopy.FieldByName('id_karty').AsInteger:=StrToInt(wczytana_karta);
end;

procedure TAOknoGl.AutoRunTimer(Sender: TObject);
Var
  plik : TextFile;
  plik_konfig, linia : String;
  czy_polaczono_db: Boolean;
  poz: Integer;
  serwer: string;
  baza: string;
  user: string;
  password: string;
  systemStart: string;
begin
 AutoRun.Enabled:=False;
 plik_konfig:=ExtractFilePath(Application.ExeName)+'Dane\konfiguracja.txt';
 if FileExists(plik_konfig)=True then
  Begin
   AssignFile(plik,plik_konfig);
   Reset(plik);
    Repeat
     Readln(plik,linia);
     if Pos('[server]',linia)>0 then
      Begin
       poz:=Pos('=',linia); Delete(linia,1,poz);
       serwer:=Trim(linia);
      End;
     if Pos('[baza]',linia)>0 then
      Begin
       poz:=Pos('=',linia); Delete(linia,1,poz);
       baza:=Trim(linia);
      End;
     if Pos('[login]',linia)>0 then
      Begin
       poz:=Pos('=',linia); Delete(linia,1,poz);
       user:=Trim(linia);
      End;
     if Pos('[haslo]',linia)>0 then
      Begin
       poz:=Pos('=',linia); Delete(linia,1,poz);
       password:=Trim(linia);
      End;
    Until eof(plik);
   CloseFile(plik);

   ADOConnection.Close;
   ADOConnection.ConnectionString:='Provider=SQLOLEDB.1;Password='+password+';Persist Security Info=True;User ID='+user+';Initial Catalog='+baza+';Data Source='+serwer;

   czy_polaczono_db := False;
    try
     try
      ADOConnection.Open();
      czy_polaczono_db:=True;
     finally
     end;
    except
    end;

    Application.ProcessMessages;
    if (czy_polaczono_db=True) then
     Begin
      //Jeśli wszystko dobrze, to uruchamiam przeszukiwanie
      StatusBar1.Panels[1].Text:='Poprawnie połączono z bazą danych KSOP Analytics';
      PageControl1.Visible:=True;
      Szukaj_karty;
      Wypelnij_liste_izotopow;
      btn_nowaClick(Self);
     end
    else ShowMessage('Nie można podłączyć się do serwera bazy danych aplikacji!');
  End
 else
  Begin
   ShowMessage('Brak pliku konfiguracyjnego!');
  End;
end;

procedure TAOknoGl.btn_czyscClick(Sender: TObject);
begin
 edt_szukaj.Clear;
 Szukaj_karty;
end;

procedure TAOknoGl.btn_nowaClick(Sender: TObject);
begin
 edt_objetosc.Text:='0';
 edt_waga.Text:='0';
 edt_zrodel.Text:='0';
 gbx_karta.Caption:='Nowa karta:';
 btn_zapisz_karte.Caption:='dodaj kartę do bazy';
 btn_zapisz_karte.Enabled:=False;
 gbx_izotopy.Visible:=False;
 edt_numer.SetFocus;
end;

procedure TAOknoGl.btn_zapisz_karteClick(Sender: TObject);
Var
 numer : String;
  id_karty: string;
begin
 numer:=Trim(edt_numer.Text);
 if btn_zapisz_karte.Caption='dodaj kartę do bazy' then
  Begin
   if Sprawdz_czy_karta_juz_istnieje(numer)=False then
    Begin
     Dodaj_nowa_karte_do_bazy;
     edt_szukaj.Text:=numer;
     Szukaj_karty;
     id_karty:=ADOKarty.FieldByName('id_karty').AsString;
     if id_karty<>'' then Wczytaj_karte(id_karty);
    End
   else ShowMessage('BŁĄD!!!'+#13+'Karta o numerze: '+numer+' już istnieje w bazie danych!');
  End
 else
  Begin
   //to znaczy, że edytujemy już istniejącą kartę
   Aktualizuj_dane_karty;
   edt_szukaj.Text:=numer;
   Szukaj_karty;
   id_karty:=ADOKarty.FieldByName('id_karty').AsString;
   if id_karty<>'' then Wczytaj_karte(id_karty);
  End;
end;

procedure TAOknoGl.cmb_rodzajChange(Sender: TObject);
Var
 rodzaj : String;
begin
 btn_zapisz_karte.Enabled:=True;
 rodzaj:=Trim(cmb_rodzaj.Text);
 if (rodzaj='Beczka barytowa') and (edt_objetosc.Text='0') then
  Begin
   edt_objetosc.Text:='0,2';
   cmb_kategoria.ItemIndex:=cmb_kategoria.Items.IndexOf('niskoaktywne');
  End;
 if (rodzaj='Bęben 200') and (edt_objetosc.Text='0') then
  Begin
   edt_objetosc.Text:='0,2';
   cmb_kategoria.ItemIndex:=cmb_kategoria.Items.IndexOf('niskoaktywne');
  End;
end;

procedure TAOknoGl.FormCreate(Sender: TObject);
begin
 AOknoGl.Caption:='KSOP Analytics - by FX Systems Piotr Daszewski - Wersja: '+wersja;
 PageControl1.Visible:=False;
 dtp_dostawa.Date:=Date;
 dtp_pomiar.DateTime:=Date;

 edt_numer.Clear;
 edt_objetosc.Text:='0';
 edt_waga.Text:='0';
 edt_zrodel.Text:='0';
 edt_symbol.Clear;
 cmb_kategoria.ItemIndex:=cmb_kategoria.Items.IndexOf('średnioaktywne');
 cmb_podkategoria.ItemIndex:=cmb_podkategoria.Items.IndexOf('krótkożyciowe');
end;

procedure TAOknoGl.FormShow(Sender: TObject);
begin
 AutoRun.Enabled:=True;
end;

procedure TAOknoGl.Zamknij1Click(Sender: TObject);
begin
 Close;
end;

procedure TAOknoGl.grd_kartyDblClick(Sender: TObject);
var
  id_karty: string;
begin
 if ADOKarty.Active then
  Begin
   id_karty:=ADOKarty.FieldByName('id_karty').AsString;
   if id_karty<>'' then Wczytaj_karte(id_karty);
  End;
end;

procedure TAOknoGl.Oprogramie1Click(Sender: TObject);
begin
 MessageBox(handle,PWideChar('KSOP Analytics'
 +#13+#13+'Wersja oprogramowania: '+wersja
 +#13+'Program opracowany przez firmę FX Systems Piotr Daszewski'
 +#13+'Różan ''2016 - ''2017'
 +#13+'Wszystkie prawa zastrzeżone!')
 ,PWideChar('Informacja o programie'),MB_OK+MB_ICONINFORMATION);
end;

procedure TAOknoGl.DBGrid2CellClick(Column: TColumn);
begin
if (Column.PickList.Count > 0) and (DBGrid2.SelectedField.AsString='') then
 begin
  keybd_event(VK_F2,0,0,0);
  keybd_event(VK_F2,0,KEYEVENTF_KEYUP,0);
  keybd_event(VK_MENU,0,0,0);
  keybd_event(VK_DOWN,0,0,0);
  keybd_event(VK_DOWN,0,KEYEVENTF_KEYUP,0);
  keybd_event(VK_MENU,0,KEYEVENTF_KEYUP,0);
 end;
end;

procedure TAOknoGl.DBGrid2ColEnter(Sender: TObject);
begin
 if (DBGrid2.SelectedField.FieldName = 'izotop') and (DBGrid2.SelectedField.AsString='') then
 Begin
  keybd_event(VK_F2,0,0,0);
  keybd_event(VK_F2,0,KEYEVENTF_KEYUP,0);
  keybd_event(VK_MENU,0,0,0);
  keybd_event(VK_DOWN,0,0,0);
  keybd_event(VK_DOWN,0,KEYEVENTF_KEYUP,0);
  keybd_event(VK_MENU,0,KEYEVENTF_KEYUP,0);
 End;
end;

procedure TAOknoGl.Dodaj_nowa_karte_do_bazy;
Begin
 ADOQuery.Close;
 ADOQuery.SQL.Clear;
 ADOQuery.SQL.Add('INSERT INTO Karty');
 ADOQuery.SQL.Add('           ([numer_karty]');
 ADOQuery.SQL.Add('           ,[data_dostawy]');
 ADOQuery.SQL.Add('           ,[rodzaj_opakowania]');
 ADOQuery.SQL.Add('           ,[symbol_opakowania]');
 ADOQuery.SQL.Add('           ,[objetosc]');
 ADOQuery.SQL.Add('           ,[waga]');
 ADOQuery.SQL.Add('           ,[ilosc_zrodel]');
 ADOQuery.SQL.Add('           ,[kategoria]');
 ADOQuery.SQL.Add('           ,[podkategoria]');
 ADOQuery.SQL.Add('           ,[data_pomiaru_aktywnosci])');
 ADOQuery.SQL.Add('     VALUES');
 ADOQuery.SQL.Add('           ('''+Trim(edt_numer.Text)+'''');
 ADOQuery.SQL.Add('           ,'''+DateToStr(dtp_dostawa.Date)+'''');
 ADOQuery.SQL.Add('           ,'''+Trim(cmb_rodzaj.Text)+'''');
 ADOQuery.SQL.Add('           ,'''+Trim(edt_symbol.Text)+'''');
 ADOQuery.SQL.Add('           ,'+StringReplace(Trim(edt_objetosc.Text),',','.',[rfReplaceAll]) );
 ADOQuery.SQL.Add('           ,'+StringReplace(Trim(edt_waga.Text),',','.',[rfReplaceAll]) );
 ADOQuery.SQL.Add('           ,'''+Trim(edt_zrodel.Text)+'''');
 ADOQuery.SQL.Add('           ,'''+Trim(cmb_kategoria.Text)+'''');
 ADOQuery.SQL.Add('           ,'''+Trim(cmb_podkategoria.Text)+'''');
 ADOQuery.SQL.Add('           ,'''+DateToStr(dtp_pomiar.Date)+''')');
 ADOQuery.ExecSQL;
End;

procedure TAOknoGl.ADOIzotopyNewRecord(DataSet: TDataSet);
begin
 ADOIzotopy.FieldByName('jednostka_akt_poczatkowej').AsString:='MBq';
end;

procedure TAOknoGl.Aktualizuj_dane_karty;
Begin
 ADOQuery.Close;
 ADOQuery.SQL.Clear;
 ADOQuery.SQL.Add('UPDATE Karty');
 ADOQuery.SQL.Add('   SET [numer_karty] = '''+Trim(edt_numer.Text)+'''');
 ADOQuery.SQL.Add('      ,[data_dostawy] = '''+DateToStr(dtp_dostawa.Date)+'''');
 ADOQuery.SQL.Add('      ,[rodzaj_opakowania] = '''+Trim(cmb_rodzaj.Text)+'''');
 ADOQuery.SQL.Add('      ,[symbol_opakowania] = '''+Trim(edt_symbol.Text)+'''');
 ADOQuery.SQL.Add('      ,[objetosc] = '+StringReplace(Trim(edt_objetosc.Text),',','.',[rfReplaceAll]));
 ADOQuery.SQL.Add('      ,[waga] = '+StringReplace(Trim(edt_waga.Text),',','.',[rfReplaceAll]));
 ADOQuery.SQL.Add('      ,[ilosc_zrodel] = '''+Trim(edt_zrodel.Text)+'''');
 ADOQuery.SQL.Add('      ,[kategoria] = '''+Trim(cmb_kategoria.Text)+'''');
 ADOQuery.SQL.Add('      ,[podkategoria] = '''+Trim(cmb_podkategoria.Text)+'''');
 ADOQuery.SQL.Add('      ,[data_pomiaru_aktywnosci] = '''+DateToStr(dtp_pomiar.Date)+'''');
 ADOQuery.SQL.Add(' WHERE id_karty='+wczytana_karta);
 ADOQuery.ExecSQL;
End;

procedure TAOknoGl.Wczytaj_karte(id_karty : String);

Begin
 ADOQuery.Close;
 ADOQuery.SQL.Clear;
 ADOQuery.SQL.Add('SELECT * FROM Karty WHERE id_karty='+id_karty);
 ADOQuery.Open;
  gbx_karta.Caption   :='Wczytana karta o numerze index: '+id_karty;
  wczytana_karta      := id_karty;
  edt_numer.Text      :=ADOQuery.FieldByName('numer_karty').AsString;
  dtp_dostawa.Date    :=StrToDate(ADOQuery.FieldByName('data_dostawy').AsString);
  dtp_pomiar.Date     :=StrToDate(ADOQuery.FieldByName('data_pomiaru_aktywnosci').AsString);
  cmb_rodzaj.ItemIndex:=cmb_rodzaj.Items.IndexOf(ADOQuery.FieldByName('rodzaj_opakowania').AsString);
  edt_symbol.Text     :=ADOQuery.FieldByName('symbol_opakowania').AsString;
  edt_objetosc.Text   :=ADOQuery.FieldByName('objetosc').AsString;
  edt_waga.Text       :=ADOQuery.FieldByName('waga').AsString;
  edt_zrodel.Text     :=ADOQuery.FieldByName('ilosc_zrodel').AsString;
  cmb_kategoria.ItemIndex   :=cmb_kategoria.Items.IndexOf(ADOQuery.FieldByName('kategoria').AsString);
  cmb_podkategoria.ItemIndex:=cmb_podkategoria.Items.IndexOf(ADOQuery.FieldByName('podkategoria').AsString);
 ADOQuery.Close;
 btn_zapisz_karte.Caption:='zapisz zmiany';
 btn_zapisz_karte.Enabled:=False;
 gbx_izotopy.Visible:=True;
 ADOIzotopy.Close;
 ADOIzotopy.CommandText:='SELECT * FROM Karty_izotopy WHERE id_karty='+wczytana_karta+' ORDER BY id_poz_izotopu ASC';
 ADOIzotopy.Open;
End;

procedure TAOknoGl.edt_numerChange(Sender: TObject);
begin
 btn_zapisz_karte.Enabled:=True;
end;

procedure TAOknoGl.edt_objetoscKeyPress(Sender: TObject; var Key: Char);
begin
 if not (Key IN [#8, ',', '.', '0'..'9']) then Key:=#0;
 if Key='.' then Key:=',';
end;

procedure TAOknoGl.edt_symbolChange(Sender: TObject);
Var
 symbol : String;
begin
 btn_zapisz_karte.Enabled:=True;
 symbol:=Trim(edt_symbol.Text);
 if symbol='PT-3' then
  Begin
   edt_objetosc.Text:='0,003';
   edt_waga.Text:='3';
   edt_zrodel.Text:='5';
   cmb_kategoria.ItemIndex:=cmb_kategoria.Items.IndexOf('średnioaktywne');
  End;
 if symbol='PT-2' then
  Begin
   edt_objetosc.Text:='0,006';
   edt_waga.Text:='6';
   edt_zrodel.Text:='10';
   cmb_kategoria.ItemIndex:=cmb_kategoria.Items.IndexOf('średnioaktywne');
  End;
end;

procedure TAOknoGl.edt_szukajKeyPress(Sender: TObject; var Key: Char);
begin
 if ORD(Key)=13 then Szukaj_karty;
end;

procedure TAOknoGl.Szukaj_karty;
var
  szuk: string;
  SQL: string;
begin
 szuk:=Trim(edt_szukaj.Text);
 if szuk<>'' then SQL:='SELECT id_karty, numer_karty, data_dostawy FROM Karty WHERE numer_karty like ''%'+szuk+'%'' ORDER BY data_dostawy, id_karty DESC'
 else SQL:='SELECT id_karty, numer_karty, data_dostawy FROM Karty ORDER BY data_dostawy DESC, id_karty DESC';
 ADOKarty.Close;
 ADOKarty.CommandText:=SQL;
 ADOKarty.Open;
end;

function TAOknoGl.Sprawdz_czy_karta_juz_istnieje(numer_karty : String): Boolean;
var
  id_karty: string;
Begin
 ADOQuery.Close;
 ADOQuery.SQL.Text:='SELECT id_karty FROM Karty WHERE numer_karty='''+numer_karty+'''';
 ADOQuery.Open;
  id_karty:=ADOQuery.FieldByName('id_karty').AsString;
 ADOQuery.Close;
 if id_karty='' then Sprawdz_czy_karta_juz_istnieje:=False
 else Sprawdz_czy_karta_juz_istnieje:=True;
End;

procedure TAOknoGl.Wypelnij_liste_izotopow;
var
  izotop: string;
Begin
 DBGrid2.Columns[0].PickList.Clear;
 ADOQuery.Close;
 ADOQuery.SQL.Text:='SELECT DISTINCT izotop FROM Okresy_polrozpadu GROUP BY izotop ORDER BY izotop ASC';
 ADOQuery.Open;
 ADOQuery.First;
 Repeat
  izotop:=ADOQuery.FieldByName('izotop').AsString;
  DBGrid2.Columns[0].PickList.Add(izotop);
  ADOQuery.Next;
 Until ADOQuery.Eof;
 ADOQuery.Close;
End;

end.
