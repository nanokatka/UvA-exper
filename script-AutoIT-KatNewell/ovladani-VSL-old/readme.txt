PROGRAM NA OVLADANIE VSL/SES

za 1.exe je potrebne zadat 4 parametre:
1. string - co sa meria (vsl alebo ses), ulozene data budu mat tuto predponu
2. maximalna dlzka pruzku v mikrometroch, maximalna hodnota je 9999
3. krok v mikrometroch, s ktorym sa bude menit dlzka pruzku
4. pocet akumulacii pre urcenie doby merania. Predpoklada sa, ze CCD je trigrovana
   externe laserom s frekvenciou 10Hz, a "gate width + gate delay" < 0.1 s

Na zaciatku je potrebne nastavit motorek do polohy 100, pomocou RESET TO a okno 
motorku dat do defaultnej polohy pomocou tlacitka REFRESH.
V ovladani CCD musi byt nastaveny adresar, do ktoreho sa budu ukladat data.