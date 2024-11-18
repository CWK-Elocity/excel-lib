# Dokumentacja do obsługi pliku Excel dla tworzenia struktury danych

---

## Wymagane rzeczy w pliku:

1. **Sekcje zapisane wielkimi literami w pierwszej kolumnie:**
   - Każda sekcja w pliku musi być wyraźnie oznaczona i zapisana wielkimi literami w pierwszej kolumnie.
   - Domyślne sekcje (ich nazwy można zmienić w pliku `env.py`):
     - `STACJA ŁADOWANIA - DANE`:
       - Kluczowa sekcja rozdzielająca dane przejęcia od danych poszczególnych stacji ładowania.
     - `OSOBA KONTAKTOWA - EKSPLOATACJA STACJI`:
       - Sekcja zawierająca dane kontaktowe dla eksploatacji stacji.
     - `OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA`:
       - Sekcja definiująca osobę odpowiedzialną za przejęcie stacji ze strony klienta.

2. **Struktura pliku Excel:**
   - **Pierwsza kolumna:** Zawiera numery lub nagłówki sekcji.
   - **Druga kolumna:** Klucze danych (np. "Imię i nazwisko", "Numer telefonu").
   - **Trzecia i kolejne kolumny:** Dane dotyczące poszczególnych stacji.

---

## Założenia i ograniczenia:

- **Sekcje wielkimi literami:**
  - Wartości w pierwszej kolumnie, które są zapisane wielkimi literami (`isupper()`), oznaczają początek nowych sekcji.
- **Spójność danych:**
  - W tej funkcji **nie zakładamy spójności danych** między kolumnami (np. dla kontaktów lub odpowiedzialności). 
  - Jeśli dane są niespójne, każda kolumna (stacja) może mieć własne dane.

---

