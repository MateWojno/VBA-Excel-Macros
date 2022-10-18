<div id=about>
<h1 align=center>Simple Coding in Excel</h1>
<p>author @MateWojno, mateusz.k.wojno@gmail.com <br>Start   17-10-2022<br>End     [?]</p>
</div>
<div id=toc> 
<h1 align=center>Table of content [PL]</h1>
  <ul>
        <li><h2>easy to read and understand user manual (how to install, run and operate - with pictures/screenshots)</h2></li>
        <li><h2>skrocic README</h2></li>
        <li><h2>zrobic odwolania, navbary, progress bary</h2></li>
        <li><h2>PR - zapraszam</h2></li>
        <li><h2>na koniec przetlumaczyc na [ENG]</h2></li>
  </ul>
</div>

<div id="res"> 
<h1 align=center>Resources:</h1>
<ul>
<li><a href="https://www.wallstreetmojo.com/vba-rename-sheet/">VBA coding</a></li>
<li><a href="https://www.wallstreetmojo.com/macros-in-excel/">Macros in Excel</a></li>
<li><a href="https://file.org/extension/bas#:~:text=BASIC%20is%20a%20programming%20language%20that%20was%20created,language%2C%20it%20is%20saved%20with%20the.bas%20file%20extension.
">.bas file extension</a></li>
<li><a href="https://www.wallstreetprep.com/self-study-programs/the-ultimate-excel-vba-course/">Paid VBA course</a></li>
<li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/resources/power-query-differences">About Power Query</a></li>
<li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences">Differences between VBA Macros and Office Scripts (online)</a></li>
<li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/">Office Scripts documentation</a></li>
<li><a href="https://en.wikipedia.org/wiki/Microsoft_Access">About MS Access</a></li>
<li><a href="https://www.lifewire.com/mdb-file-2621974">What Is an MDB File?</a></li>
<li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons?source=recommendations">About scripted buttons in Microsoft Excel Desktop App</a></li>
</ul>
</div>

###     `Tools/extensions required:`
-       Microsoft Excel 2019 (pro recommended);
-       Power Query (built in);
-       Visual Studio Code (recommended for Devs) + extensions: XVBA - Live Server VBA, VBA v0.6.0 serkonda7, vba-snippets Scott Spence;
-       Dedicated App built by me (optional)
-       for old .mdb files you need MS Access 2010

##      `#1/    Reading database file     -   algorithm`

###     `Input data:`
*               MS Access database file:
-               .mdb/.accdb;
-               
###     `Description:`
-       [PL]
            1/  start
            2/  pobierz dane w programie excel z okreslonej sciezki;
            3/  zaimportuj te dane tworzac nowy arkusz;
            4/  nazwij nowy arkusz "wdb", tak aby zawsze mozna bylo sie do niego odwolac;
            5/  ustaw "wdb" jako aktywny arkusz;
            6/  zapisz log o alternatywnej tresci 1 - "%date% %username% udalo sie" OR 0 - "%date% %username% wystapil blad"
            4/  stop

###     `Output data:`
-       [PL]
            1/  tabela w Excelu, ktora wymaga filtrowania wynikow;
            2/  zawarta w nowym, nazwanym, aktywnym arkuszu;

##      `#2/    Data transformation in power query     -   algorithm`

###     `Input data:`
-       output from algorithm #1, .xlsm (Excel with Macros);
-       active excel sheet;

###     `Description:`
-       [PL]
        1/  start;
        2/  wczytaj dane z algorytmu #1 i zaznacz aktywny arkusz;
        3/  przetransformuj tabele wedlug okreslonego wzoru naglowkow;
        4/  zmien nazwy okreslonych naglowkow wedlug wzoru;
        5/  dopasuj kolejnosc naglowkow tabeli do wzoru;
        6/  sformatuj odpowiednio zawartosc komorek, zaokraglenie do 2 cyfr znaczacych;
        7/  zapisz log o alternatywnej tresci 1 - "%date% %username% udalo sie" OR 0 - "%date% %username% wystapil blad";
        8/  stop;

###     `Output data:`
-       table in sheet, format .xlsx or .xlsm (prefered)

##      `#3/    Refresh loop     -   algorithm`

###     `input data:`
-       [PL]
            1/  output z algorytmu #3;
            2/  aktywny arkusz MS Excel po transformacji o stalej nazwie "wdb";

###     `description:`

#####   *short info why this app needs it*
-       [PL]
            VBA macra w MS Excel niestety maja jeden zasadniczny problem;
            po nacisnieciu przycisku, w ktorym zapisane jest nagrane macro (wybierz plik, zaimportuj, zmien nazwe, przeksztalc - czyli zawierajacy dzialanie algorytmow #1 i #2) nie mozna ponownie uzyc tego przycisku z dwoch powodow:
*           1/  baza danych jest aktualnie uzywana;
*           2/  arkusz o danej nazwie jest aktualnie uzywany;

#####   *weak solution of this issue*
-       [PL]
            Nalezy recznie:
*           1/ usunac polaczenie z baza danych;
*           2/ usunac aktywny arkusz;

#####   *why i do it manually*
-       [PL]
            Poniewaz:
*           1/ adresowanie aktywnego arkusza, podczas nagrania makra odwoluje sie do konkretnego numeru aktywnego arkusza; 
*           2/ w procesie tworzenia nowego arkusza (nawet po zmianie jego nazwy) numer (index) tego arkusza zostaje powiekszony o 1 (++i zamiast i++    -   inkrementacja)  -   blad rekurencji, poniewaz nie mam kontroli nad calkowitym jej przebiegiem;
*           3/ makro (algorithm), ktore kasuje polaczenie z baza danych oraz wybrany, aktywny arkusz nie nadaje sie do wielokrotnego uzycia, poniewaz index arkusza (dane wejsciowe) nie aktualizuja sie odpowiednio, w ten sposob makro wciaz oczekuje starych danych - tzn. dziala ale tylko raz;
#####   *strong solution of this issue*
-       [PL]
            Aby rozwiazac powyzsze problemy nalezy:
*           1/ zmienic sposob adresowania aktywnego arkusza w procesie algorytmu #1 z rekurencyjnego n = n++ na iteracyjny n = ++n albo calkowicie zmienic sposob adresowania bez iteracji ani rekurencji
            *(o ile dobrze to rozumiem, zalezy mi na tym, zeby mozna bylo odpowiednio odwolac sie do wartosci podstawowej n sprzed iteracji/rekurencji, ewentualnie zupelnie nalezy czyscic tmp/cache tej aplikacji, tak aby nie zapamietywala, ze wczesniej byl jakikolwiek arkusz utworzony)*
*           2/ podobnie nalezy postapic z zerwaniem polaczenia z baza danych i oczyszczeniem pamieci komputera o  tym, ze jakakolwiek baza byla uprzednio polaczona;


####    `list of steps:`
        [PL]
            1/  start;
            2/  usun polaczenie z baza danych i arkusz "wdb";
            3/  wyczysc pamiec cache albo tmp programu na ten temat;
            4/  zapisz log o alternatywnej tresci 1 - "%date% %username% udalo sie" OR 0 - "%date% %username% wystapil blad";
            5/  stop;

###     `output data:`
-       [PL]
            1/ ten algorytm polega na czyszczeniu pliku i pamieci, nie generuje outputu;

##      `#4/    Automatization of this loop     -   algorithm`

###     `input data:`

###     `description:`

###     `output data:`

###     `API (interface):`

##      `#5/    Debug/Bug tracker      -       for app or addon:`
-       [PL]
            1/          zbiera logi i ewentualne bledy;
            2/          pozwala zglaszac bledy uzytkownikowi wraz z ich opisem;
            3/          przesyla wiadomosci o bledach na adres tworcy za pomoca programu pocztowego; 

#       `API (interface):`
-       [PL]
            #1/         przycisk w arkuszu [main] <Data-Fetch>;
            #2/         przycisk w arkuszu [main] <Data-Transform>;
            #3/         przycisk w arkuszu [main] <Data-Clean>;
            #4/         przycisk w arkuszu [main] <Auto>;
            #5/         przycisk w arkuszu [main] <Debug>;







