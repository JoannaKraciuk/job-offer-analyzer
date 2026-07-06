# Job Offer Analyzer

Minimalny pierwszy krok narzędzia AI do śledzenia ofert pracy.

Aktualny zakres:

- trzyma roboczą kopię pliku Excel w `data/`
- dodaje jedną testową ofertę do arkusza `Oferty`
- dodaje jeden wpis kontrolny do arkusza `Historia_Sprawdzeń`
- automatycznie generuje kolejny identyfikator `JOB-XXX`
- pobiera treść oferty z linku i uzupełnia formularz danymi znalezionymi na stronie

Uruchomienie:

```powershell
python app.py
```

Opcjonalne dane testowe:

```powershell
python app.py --company "Example Corp" --title "Junior QA Engineer" --link "https://example.com/job"
```

Uruchomienie UI w przeglądarce:

```powershell
streamlit run ui.py
```

W UI wklej link do oferty i kliknij `Pobierz dane`. Pola formularza zostaną uzupełnione automatycznie, ale przed zapisem nadal można je ręcznie poprawić.

Kliknij `Sprawdź oferty`, żeby aplikacja przeszła po linkach zapisanych w arkuszu `Oferty`, zaktualizowała `Dostępność`, `Ostatnio sprawdzono`, `Historia_Sprawdzeń` oraz `Dashboard`.

Kliknij `Uzupełnij stawki`, żeby aplikacja ponownie przeszła po zapisanych linkach i uzupełniła kolumny wynagrodzenia dla istniejących rekordów.

## Wynagrodzenie

Aplikacja próbuje wykryć widełki najpierw z danych strukturalnych oferty, np. `baseSalary`, a jeśli ich nie ma, z treści strony. Zapisuje:

- oryginalny fragment ze stawką
- walutę
- stawkę min/max
- okres: godzinowo, miesięcznie albo rocznie
- brutto/netto
- kurs waluty do PLN
- przeliczenie na PLN miesięcznie i godzinowo
- założenia przeliczenia

MVP nie liczy podatków ani realnego netto. Do porównań używa założenia `160 h/mies.` i kursu NBP tabela A dla EUR/USD.

## Instalacja

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Uruchomienie

```powershell
.\.venv\Scripts\python.exe -m streamlit run ui.py
```

## Dane prywatne

Pliki `.xlsx` w katalogu `data/` są ignorowane przez Git. Nie commituj realnej bazy ofert, backupów ani prywatnych linków aplikacyjnych.
