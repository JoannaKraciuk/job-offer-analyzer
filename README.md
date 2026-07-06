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
