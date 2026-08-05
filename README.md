# Měření činností

Mobilní aplikace pro skenery:

- výběr pracovníka,
- výběr činnosti,
- START / END,
- uložení do Supabase,
- export posledních 24 hodin do Excelu.

## Soubory

- `app.py` – aplikace
- `database.sql` – vytvoření tabulky a pravidel v Supabase
- `requirements.txt` – Python knihovny
- `.streamlit/secrets.toml.example` – vzor připojení
- `.gitignore` – chrání skutečný soubor s hesly před nahráním na GitHub

## Zprovoznění

1. Vytvoř projekt na Supabase.
2. Otevři SQL Editor, vlož celý obsah `database.sql` a spusť jej.
3. V Supabase najdi Project URL a anon/public key.
4. Nahraj projekt na GitHub.
5. Na Streamlit Community Cloud vytvoř aplikaci z `app.py`.
6. V nastavení aplikace otevři Secrets a vlož:

```toml
[supabase]
url = "https://TVUJ-PROJEKT.supabase.co"
key = "TVUJ_SUPABASE_ANON_KEY"
```

## Poznámka k exportu

Aktuální verze vytvoří Excel po stisknutí tlačítka a zahrne záznamy
za posledních 24 hodin. Automatické každodenní odesílání nebo ukládání
souboru vyžaduje samostatný naplánovaný proces.
