Slozka se skripty (wingate_skript.r a export.rmd) musi obsahovat nasledujici podslozky (nazvy bez diakritiky, tak jak jsou psany):
antropometrie - obsahuje xls soubory s antropometrii (jeden hromadný nebo pro každého probanda zvlášť)
databaze - vysledna excel databaze se vsemi predchozimi merenimi
reporty - vysledne pdf reporty po spusteni skriptu
vymazat - docasna data, ktera je treba vymazat po dokonceni skriptu
vysledky - obsahuje excel export ze soucasneho mereni
wingate - txt exporty z Wattbike

SLOZKY ANTROPOMETRIE, LAKTAT, WINGATE MUSI OBSAHOVAT POUZE DATA Z AKTUALNIHO MERENI - ostatni presunte do slozky jiz probehla mereni. 
SLOZKY REPORTY, VYMAZAT by mely byt pred spustenim skriptu prazdne. 

Slozky ANTROPOMETRIE a WINGATE musi obsahovat soubory se shodnymi nazvy (id), stejne tak id musi odpovidat v excelu laktat.xls, jinak se soubory nesparuji
a skript zobrazi upozorneni (viz krok 8 nize).

Postup:
1) rozkliknout soubor wingate_skript
2) na 5. radku skriptu zkontrolovat cestu ke slozce se skriptem ve formatu "C:/Users/<Uzivatel>/Documents/Wingate skript"
3) spustit pomoci tlacitka <source> vpravo nahore podokna s kodem skriptu
4) vlevo dole (dale konzole) se objevi dotaz Sport?: kliknete vedle dotazu a napiste sport testovaných, potvdte klavesou enter
5) to same opakujte pro dotaz Tým?:
6) skript kontroluje kompatibilitu jmen souboru (wingate, antropometrie a laktat-id ve sloupci 1)
7) pokud jsou id v poradku, skript pokracuje v generování reportu - pokracujte na krok 13)
8) pokud jsou id spatne, skript zobrazi dole v konzoli upozorneni, ktere id nesouhlasi a v jakem souboru + dotaz o pokracovani v exportu
9) napiste dolu do konzole N a stisknete klavesu enter, skript bude prerusen
10) opravte jmena souboru
11) spustte skript znovu pomoci tlacitka <source>
12) opakujte krok 4) a 5)
13) skript generuje reporty do slozky <reporty>
14) skript zobrazi dotaz k ulozeni souboru do databaze, napiste do konzole A nebo N a potvrdte enter
15) pokud A - databaze nahraje posledni ulozeny databazovy soubor a prida k nemu soucasne mereni, nasledne vyexportuje novy soubor. 
16) vymazte vse ve slozce vymazat

V pripade dotazu volejte 604774455, Dominik Kolinger




