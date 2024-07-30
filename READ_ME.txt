Slozka se skripty (Wingate_Spiro_radarchart.r a export.rmd) musi obsahovat nasledujici podslozky (nazvy bez diakritiky, tak jak jsou psany):
antropometrie - obsahuje .xls soubory s antropometrii (jeden hromadný nebo pro každého probanda zvlášť), v pripade vyhodnoceni skoku soubor obsahuje sloupec "SJ" s ciselnymi hodnotami
databaze - vysledna excel databaze se vsemi predchozimi merenimi
reporty - vysledne pdf reporty po spusteni skriptu
vymazat - docasna data, ktera je treba vymazat po dokonceni skriptu
vysledky - obsahuje excel export ze soucasneho mereni
wingate - .txt exporty z Wattbike (aktuální měření), složky pro porovnání můžou být umístěny kdekoliv (skript spustí dotaz na cestu k souboru)
spiro - .xlsx exporty ze spiroergometrie


SLOZKY REPORTY, VYMAZAT by mely byt pred spustenim skriptu prazdne. 

Slozky ANTROPOMETRIE a WINGATE a pripadne SPIRO musi obsahovat soubory se shodnymi nazvy (id), stejne tak id musi odpovidat v excelu laktat.xls, jinak se soubory nesparuji
a skript zobrazi upozorneni (viz krok 8 nize).

Postup:
1) rozkliknout soubor Wingate_Spiro_radarchart
2) na 3. radku skriptu zkontrolovat cestu ke slozce se skriptem ve spravnem formatu (např. "C:/Users/<Uzivatel>/Documents/Wingate+Spiro")
3) spustit pomoci tlacitka <source> vpravo nahore podokna s kodem skriptu
4) objevi se dotaz na spiro report, pokud date ANO, soubory ze slozky spiro budou zahrnuty do reportu
5) dalsi dotaz je pro srovnavaci wingate soubory, lze pridat dve srovnani
6) nasledujici dotaz je pro pridani nejstarsiho srovnani do grafu, pokud date NE, ve vyslednem grafu nebude zahrnuta krivka nejstarsiho srovnani
7) nasledne skript kontroluje kompatibilitu vsech souboru (pojmenovani a pocet), v pripade chyby je zobrazena chybova hlaska a popis v konzoli 
8) pokud u chybove hlasky zvolite pokracovat v exportu, chybna data nebudou vyhodnocena, pripadne budou vyhodnocena castecne
9) pokud zvolite ne, skript se zastavi, opravte soubory a zacnete znovu stisknutim tlačítka <source>
7) pokud jsou id a pocty v poradku, skript pokracuje v generování reportu
10) skript generuje reporty do slozky <reporty>
11) skript zobrazi dotaz k ulozeni souboru do databaze, napiste do konzole A nebo N a potvrdte enter
12) pokud A - databaze nahraje posledni ulozeny databazovy soubor a prida k nemu soucasne mereni, nasledne vyexportuje novy soubor. 
13) vymazte vse ve slozce vymazat

do.kolinger@gmai*.com




