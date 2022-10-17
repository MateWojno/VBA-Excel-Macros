#       `VBA/TS     -   scripting macros in Excel`
-       author @MateWojno, mateusz.k.wojno@gmail.com
-       Start   17-10-2022
-       End     [?]

##      `Table of content`
-       ToDo:
-           style this file better and make some links/button from Table of content to headers;
-           make some navbar or something;
-           translate from pl to eng properly;
-           make easy progress bar or something to track progress;
-           testing and progress bar;
-           issues/bugs tracker;
-           easy to read and understand user manual (how to install, run and operate) with pictures/screenshots;

##      `Resources:`

###     `About Power Query`
*   https://learn.microsoft.com/en-us/office/dev/scripts/resources/power-query-differences

###     `Differences between VBA Macros and Office Scripts`
*   https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences

###     `Office Scripts documentation`
*   https://learn.microsoft.com/en-us/office/dev/scripts/

###     `About MS Access     -   C++ back-end`
*   https://en.wikipedia.org/wiki/Microsoft_Access

#### `Most important:`
-   File extensions
-   Microsoft Access saves information under the following file formats:
-   Current formats
-   File format	Extension
-   Access Blank Project Template	.adn
-   Access Database (2007 and later)	.accdb
-   Access Database Runtime (2007 and later)	.accdr
-   Access Database Template (2007 and later)	.accdt
-   Access Add-In (2007 and later)	.accda
-   Access Workgroup, database for user-level security.	.mdw
-   Protected Access Database, with compiled VBA and macros (2007 and- later)	.accde
-   Windows Shortcut: Access Macro	.mam
-   Windows Shortcut: Access Query	.maq
-   Windows Shortcut: Access Report	.mar
-   Windows Shortcut: Access Table	.mat
-   Windows Shortcut: Access Form	.maf
-   Access lock files (associated with .accdb)	.laccdb

###     `What Is an MDB File?`
-   https://www.lifewire.com/mdb-file-2621974

*   A file with the MDB file extension is a Microsoft Access database file that literally stands for Microsoft Database. This is the default database file format used in Access 2003 and earlier, while newer versions use the ACCDB format.

*   MDB files contain database queries, tables, and more that can be used to link to and store data from other files, like XML and HTML, and applications, like Excel and SharePoint. An LDB file is sometimes seen in the same folder as an Access database file; it's an Access lock file that's temporarily stored with a shared database.

###     `About scripted buttons in Microsoft Excel Desktop App`
*   https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons?source=recommendations  

##      `What am I going to script and why:`

*           VBA     -   scripts (inside buttons)     for:
-           1/   reading database file;
-           2/   data transformation in power query;
-           3/   refresh loop;
-           4/   automatization of this loop;

###   `Why?    -   to make our Company employees more productive;`

##      `#1/   reading database file     -   algorithm`

###     `tools/extensions required:`
-       Microsoft Excel 2019 (pro recommended);
-       Power Query (built in);
-       Visual Studio Code (recommended for Devs)
-       Dedicated App built by me (optional)
####   *for old .mdb files you need MS Access 2010*

###     `input data:`
*       MS Access database file:
-       .mdb/.accdb;

###     `description:`
-       [PL]
            1/  start
            2/  pobierz dane w programie excel z okreslonej sciezki;
            3/  zaimportuj te dane tworzac nowy arkusz;
            4/  nazwij nowy arkusz "wdb", tak aby zawsze mozna bylo sie do niego odwolac;
            5/  ustaw "wdb" jako aktywny arkusz;
            6/  zapisz log o alternatywnej tresci 1 - "%date% %username% udalo sie" OR 0 - "%date% %username% wystapil blad"
            4/  stop

###     `output data:`
-       [PL]
            1/  tabela w Excelu, ktora wymaga filtrowania wynikow;
            2/  zawarta w nowym, nazwanym, aktywnym arkuszu;

###     `API (interface):`
-       [PL]
            1/  przycisk w arkuszu [main] <Data-Fetch>;
            2/  kiedys - wlasny, zewnetrzny albo wbudowany addon do Excela, dedykowany do tego rozwiazania;

###     `Debug/Bug tracker      -       for app or addon:`
-       [PL]
            1/  zbiera logi i ewentualne bledy;
            2/  pozwala zglaszac bledy uzytkownikowi wraz z ich opisem;
            3/  przesyla wiadomosci o bledach na adres tworcy za pomoca programu pocztowego;            


##      `#2/   data transformation in power query     -   algorithm`

###     `tools/extensions required:`
-       Microsoft Excel 2019 (pro recommended);
-       Power Query (built in);
-       Visual Studio Code (recommended for Devs)
-       Dedicated App built by me (optional)
####   *for old .mdb files you need MS Access 2010*

###     `input data:`
-       output from algorithm #1, .xlsm (Excel with Macros);
-       active excel sheet;

###     `description:`
-       [PL]
        1/  start;
        2/  wczytaj dane z algorytmu #1 i zaznacz aktywny arkusz;
        3/  przetransformuj tabele wedlug okreslonego wzoru naglowkow;
        4/  zmien nazwy okreslonych naglowkow wedlug wzoru;
        5/  dopasuj kolejnosc naglowkow tabeli do wzoru;
        6/  sformatuj odpowiednio zawartosc komorek, zaokraglenie do 2 cyfr znaczacych;
        7/  zapisz log o alternatywnej tresci 1 - "%date% %username% udalo sie" OR 0 - "%date% %username% wystapil blad";
        8/  stop;

###     `output data:`
-       table in sheet, format .xlsx or .xlsm (prefered)

###     `API (interface):`
-       [PL]
            1/  przycisk w arkuszu [main] <Data-Transform>;
            2/  kiedys - wlasny, zewnetrzny albo wbudowany addon do Excela, dedykowany do tego rozwiazania;

###     `Debug/Bug tracker      -       for app or addon:`
-       [PL]
            1/  zbiera logi i ewentualne bledy;
            2/  pozwala zglaszac bledy uzytkownikowi wraz z ich opisem;
            3/  przesyla wiadomosci o bledach na adres tworcy za pomoca programu pocztowego;            



##      `#3/   refresh loop     -   algorithm`

###     `tools/extensions required:`
-       Microsoft Excel 2019 (pro recommended);
-       Power Query (built in);
-       Visual Studio Code (recommended for Devs)
-       Dedicated App built by me (optional)
####   *for old .mdb files you need MS Access 2010*

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

###     `API (interface):`
-       [PL]
            1/  przycisk w arkuszu [main] <Data-Clean>;
            2/  kiedys - wlasny, zewnetrzny albo wbudowany addon do Excela, dedykowany do tego rozwiazania;


##      `#4/   automatization of this loop     -   algorithm`

###     `tools/extensions required:`
-       Microsoft Excel 2019 (pro recommended);
-       Power Query (built in);
-       Visual Studio Code (recommended for Devs)
-       Dedicated App built by me (optional)
####   *for old .mdb files you need MS Access 2010*
###     `input data:`

###     `description:`

###     `output data:`

###     `API (interface):`







