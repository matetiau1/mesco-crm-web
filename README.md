# MESco CRM Web — wdrożenie

Aplikacja webowa rozszerzająca CRM (Power Apps). Czyta i zapisuje dane bezpośrednio w SharePoint **CRM System POC** przez Microsoft Graph, po zalogowaniu kontem M365. W kodzie nie ma żadnych danych ani haseł, więc repozytorium może być publiczne.

## Co zawiera

| Moduł (menu) | Zakładki |
|---|---|
| DASHBOARD | To-Do (follow-upy, Zrobione, Dodaj notatkę, Zadzwoń), Aktywności, Raport sprzedaży (Szanse / TECSy / Tech Pool, marża, premie Q/R, NTA/NTM), Godziny TECS / PreSales, Bilans firm (z „Przelicz”) |
| TECSY | Przegląd, Do zamknięcia, Zamknięte, Kontrola cen, TP, Bez oferty, Bez notatek |
| OPPSY | Przegląd (lejek wg Etap1), W toku, Wygrane / przegrane (Status1), Alerty |
| FIRMY | Wykorzystanie TECS — support z 90 dni / 12 mies., alert „90 dni bez supportu” |
| KODY DEMO | Aktywne i kończące się, Wszystkie, Konwersja (opiekun, produkt) |

Pozostałe pozycje menu (HR, Kontakty, Kalendarz…) są wyszarzone. Po wpisaniu linku do aplikacji Power Apps w `config.js` → `powerAppsUrl` będą ją otwierać.

## Krok 1 — repozytorium GitHub

1. Na github.com utwórz nowe repozytorium, np. **`mesco-crm-web`** (Public).
2. **Add file → Upload files** i przeciągnij wszystkie pliki z tego folderu: `index.html`, `app.js`, `app.css`, `config.js`, `sw.js`, `manifest.webmanifest`, `icon.svg`, `icon-192.png`, `icon-512.png`. Na koniec **Commit changes**.
3. **Settings → Pages**: Source = *Deploy from a branch*, Branch = `main` / `(root)` → **Save**.
4. Po ok. minucie aplikacja będzie pod adresem `https://matetiau1.github.io/mesco-crm-web/`.

## Krok 2 — aplikacja w Entra ID (raz)

Portal Entra → **App registrations** → aplikacja `7277cae3-d05a-44f6-8afe-2edb6b9b6365`:

1. **Authentication → Add a platform → Single-page application** i dodaj adresy przekierowania:
   - `https://matetiau1.github.io/mesco-crm-web/`
   - `https://matetiau1.github.io/mesco-crm-web/index.html`

   Jeśli platforma SPA już istnieje (dodatek Outlook), dopisz te adresy do niej.
2. **API permissions → Add a permission → Microsoft Graph → Delegated** i zaznacz `Sites.ReadWrite.All` (`User.Read` zwykle już jest). Następnie **Grant admin consent for MESco**.

> Uprawnienie jest *delegowane*: każdy użytkownik widzi i zmienia w SharePoint tylko to, do czego ma dostęp. Aplikacja nie ma żadnych własnych uprawnień.

## Krok 3 — instalacja na PC

Otwórz adres w Edge lub Chrome i zaloguj się kontem MESco. Potem w Edge wybierz **⋯ → Aplikacje → Zainstaluj tę witrynę jako aplikację**, a w Chrome ikonę instalacji w pasku adresu. Aplikacja pojawi się w menu Start i na pasku zadań. Ten sam link możesz wysłać handlowcom.

## Aktualizacje

Podmieniasz plik w repozytorium (Upload files, ta sama nazwa), a zmiana jest widoczna po odświeżeniu. Dane są zawsze aktualne; przycisk **Odśwież** pobiera je ponownie bez przeładowania strony.

## Zasady liczenia

- **Kwartały**: rok obrotowy Ansys XI–X (Q1 = lis–sty), jak w Power Apps.
- **TECS**:
  - do zamknięcia: Przygotowanie oferty / Oferta wysłana;
  - wygrane: Zamówione / Wystawiona faktura;
  - przegrane: Rezygnacja z TECS;
  - **TP** i **Demo** są osobnymi kategoriami.
- **Wartość TECS**: suma `Oferty_TECS.SUMA` (łączone po `IdentyfikatorTECS` = `TECSGUID`); bez ofert — pole `Wartość`.
- **Kontrola cen**: `Cena_TECS` vs kolumna *TECS Fee* z `Cennik_CRM`; dla TP — *Annual Lease*.
- **Szanse sprzedaży**: status z `Status1`, etap z `Etap1`; marża 35% (New Business / B2TECS) lub 45% (New Logo); premie jak w Power Apps.
- **Notatki**: łączone po `Opps_ID` = GUID szansy; support po `IdentyfikatorCompanies`.
- **Bilans firm / Przelicz**: zapisuje do `Companies` — `WartoscTECS = (TECSh + PreSalesh) × 75`, `MarzaTECS = 33% aktywnych TECS`, szkolenia gratis z 12 mies. / kurs EUR.
