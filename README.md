# xflow-anlaegskartotek
[**Formål**](#formål) | [**Beskrivelse**](#beskrivelse) | [**Features**](#features) | [**Afhængigheder**](#afh%C3%A6ngigheder)

# Formål

Automatisering for process af indberetninger til vedligeholdelse af anlægskartotek.
Applikationen sender en e-mail med en Excel-rapport og eventuelle vedhæftede leasingkontrakter
til den konfigurerede modtager.

## Beskrivelse

1. Applikationen modtager en indberetning fra XFlow.
2. Indberetningen samles i en Exce-rapport.
3. Eventuelle leasingkontrakter vedhæftes.
4. Rapporten og filerne sendes via e-mail til den konfigurerede modtager

## Features

- Sender en Excel-rapport via e-mail.
- Vedhæfter eventuelle leasingkontrakter.
- Gør det muligt at konfigurere e-mailmodtageren via miljøvariablen `REPORT_EMAIL_RECIPIENT`.

## Afhængigheder

:cloud: | API'er, FTP-servere, CDN'er andre online services som applikationen kommunikerer med.

- **XFlow** - leverer indberetningen.
- **SMPT-server** - sender e-mailen med rapport og vedhæftede filer.

:heavy_dollar_sign: | Miljøvariabler som forventes at være populeret.

| Variabel | Beskrivelse |
|---|---|
| `FLASK_ENV` | Konfiguration af miljøet: `dev` eller `prod`. Standard er `dev`. |
| `PORT` | Port, som applikationen lytter på. Standard er `8080`. |
| `DEBUG` | Slår debug-tilstand til eller fra. Standard afhænger af `FLASK_ENV`. |
| `SMTP_SERVER` | Adresse på SMTP-serveren. |
| `SMTP_PORT` | Port på SMTP-serveren. |
| `SMTP_SENDER` | E-mailadresse, der sender rapporterne. |
| `SMTP_PASSWORD` | Adgangskode til afsenderkontoen. |
| `REPORT_EMAIL_RECIPIENT` | Modtager af rapporterne. | 

Adgangskoder og andre hemmeligheder skal leveres sikkert. 

## Udvikling

<!--
Beskrivelse af hvordan løsningen startes lokalt.
Overblik over eventuelle API endpoints som løsningen udstiller.
Notering af eventuelle hard-codede værdier, og hvordan værdierne kan ændres.
-->

### Lokal udvikling

1) Opret et virtuelt miljø:

    *Linux/macOS*:

    ```bash
    python -m venv .venv
    source .venv/bin/activate
    ```

    *Windows (PowerShell)*:

    ```powershell
    python -m venv .venv
    .\.venv\Scripts\Activate.ps1
    ```

2) Installér dependencies:

    ```bash
    pip install -r requirements.txt -r requirements-dev.txt 
    ```

3) Opret din `.env` (hvis denne mangler)

    Kopiér eksempel-filen, og redigér værdier efter behov.

    ```bash
    cp .env.example .env
    ```

    Bemærk:

    - `.env` ignoreres af git.
    - Hvis du tilføjer nye miljøvariabler, skal du også tilføje dem i `src/utils/config.py`.

4) Kør applikationen:

    ```bash
    python src/main.py
    ```

### Docker (gunicorn)

Applikationen benytter Docker via `Dockerfile`, som kører applikationen med gunicorn.

Byg og kør applikationen:

```bash
docker build -t xflow_anlaegskartotek:local .
docker run --rm -p 8080:8080 -e FLASK_ENV=prod -e PORT=8080 xflow_anlaegskartotek:local
```

Kør applikationen med compose:

```bash
docker compose up
```

### Formatering, linting, test

Organisér imports med iSort:

```bash
isort .
```

Formatér kode, så den overholder PEP 8 med Autopep8:

```bash
autopep8 .
```

Kør linting med Flake8:

```bash
flake8
```

Kør tests med Pytest:

```bash
pytest
```