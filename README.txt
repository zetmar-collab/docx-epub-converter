DOCX EPUB Converter - aplikacja desktopowa
Autor: Marek Zettel

Windows:
1. Kliknij prawym przyciskiem myszy install_windows.ps1 i wybierz "Run with PowerShell".
   Jesli system blokuje skrypt, uruchom w PowerShell:
   powershell -ExecutionPolicy Bypass -File .\install_windows.ps1
2. Instalator utworzy srodowisko .venv, zainstaluje zaleznosci i doda skrot
   "DOCX EPUB Converter" na pulpicie.
3. Do szybkiego uruchomienia bez instalacji skrotu uzyj START_EPUB_CONVERTER.bat.

macOS:
1. W Terminalu przejdz do folderu projektu.
2. Uruchom:
   chmod +x install_macos.sh START_EPUB_CONVERTER.command
   ./install_macos.sh
3. Instalator utworzy aplikacje "DOCX EPUB Converter.app" na pulpicie.
4. Do szybkiego uruchomienia bez instalacji aplikacji uzyj START_EPUB_CONVERTER.command.

Wymagania:
- Python 3.10 lub nowszy.
- EpubCheck moze wymagac zainstalowanej Javy.

Sposob pracy:
1. Uzupelnij metadane ebooka. Wymagane sa wszystkie pola oprocz ISBN.
2. Wybierz plik DOCX i okladke PNG/JPG.
3. Wybierz folder zapisu.
4. Kliknij "Konwertuj i zapisz EPUB".
5. Przed konwersja aplikacja sprawdza, czy DOCX ma poprawna strukture
   rozdzialow dla EPUB. Jesli nie, pokazuje osobne okno z lista poprawek.
6. Po konwersji aplikacja pokazuje spis tresci, podglad rozdzialow,
   wynik walidacji i przyciski otwierania EPUB/folderu.
