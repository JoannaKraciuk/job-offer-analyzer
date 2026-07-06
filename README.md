# Job Offer Tracker

Job Offer Tracker pomaga uporządkować aplikowanie na oferty pracy QA/IT. Aplikacja zbiera dane z linków do ogłoszeń, zapisuje je w jednym pliku Excel i pozwala szybko sprawdzić, które oferty wymagają kolejnej akcji.

Excel jest w tym MVP warstwą przechowywania i raportowania danych. Plik `.xlsx` działa jak lokalna baza ofert, historia sprawdzeń oraz dashboard do codziennej pracy nad aplikacjami.

## Aktualne funkcje

- pobieranie danych z URL oferty pracy
- ręczna korekta danych przed zapisem
- zapis nowych ofert do arkusza `Oferty`
- wykrywanie duplikatów po linku
- aktualizacja dostępności istniejących ofert
- uzupełnianie stawek i przeliczeń wynagrodzenia
- automatyczne wykrywanie portalu na podstawie URL
- zapisywanie historii akcji i sprawdzeń
- porządkowanie dashboardu i najbliższych działań

## Struktura danych w Excelu

Główne arkusze:

- `Oferty` - aktualna lista ofert i status procesu
- `Analiza_CV` - miejsce na analizę wymagań oferty względem profilu kandydata
- `Historia_Sprawdzen` / `Historia_Sprawdzeń` - dziennik akcji: dodanie, aktualizacja, niedostępność albo pominięcie duplikatu
- `Pytania_Formularzy` - miejsce na pytania z formularzy aplikacyjnych
- `Dashboard` - podsumowanie i najbliższe działania

Najważniejsze kolumny raportowe w `Oferty`:

- `Match Score` - liczbowy wynik dopasowania, np. `82`
- `Priority` - `HIGH`, `MEDIUM` albo `LOW`
- `Technologie` - technologie rozdzielone średnikiem, np. `Python; Playwright; SQL; API`
- `Portal` - wykryte źródło oferty, np. `Just Join IT`, `No Fluff Jobs`, `LinkedIn`, `TestDevJobs` albo `Unknown`
- `Ostatnia akcja` - ostatni etap procesu, np. `Dodano`, `Zaktualizowano`, `CV wysłane`, `HR`, `Techniczna`, `Oferta`, `Odrzucona`

## Instalacja

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Uruchomienie UI

```powershell
.\.venv\Scripts\python.exe -m streamlit run ui.py
```

Po uruchomieniu wklej link do oferty, kliknij `Pobierz dane`, popraw pola w formularzu, jeśli trzeba, i zapisz ofertę do Excela.

## Dane prywatne

Pliki `.xlsx` w katalogu `data/` są ignorowane przez Git. Nie commituj realnej bazy ofert, backupów ani prywatnych linków aplikacyjnych.

Prywatne dane profilu CV powinny znajdować się w `data/private/cv_profile.yml`. Katalog `data/private/` jest ignorowany przez Git.

## Planned Features

- CV matching
- AI analysis
- market trends dashboard
