# Kontekst badań i zadanie skryptu

## 1. Cel badań

Celem badań jest porównanie centralnej optymalizacji PSO z prostym lokalnym sterowaniem falownikami PV i magazynem energii w asymetrycznej sieci niskiego napięcia.

Badania mają odpowiedzieć na pytanie:

**Czy lokalne sterowanie falownikami PV według charakterystyki Q(U), uzupełnione skokową regułą pracy magazynu energii, może uzyskać znaczną część efektu centralnej optymalizacji PSO przy mniejszej złożoności pomiarowej i komunikacyjnej?**

Nie chodzi o ponowne porównywanie algorytmów metaheurystycznych. PSO jest tutaj wariantem referencyjnym, czyli „najlepszym technicznie” punktem odniesienia.

---

## 2. Model i założenia

Model bazowy: zmodyfikowany **IEEE European Low Voltage Test Feeder** w środowisku **DIgSILENT PowerFactory**.

Sieć musi być analizowana trójfazowo i asymetrycznie, osobno dla faz:

- L1,
- L2,
- L3.

Nie wolno sprowadzać obliczeń do modelu symetrycznego.

Scenariusz główny:

- ok. 20 kW odbiorów,
- ok. 80 kW generacji PV,
- pojedynczy punkt pracy,
- wysoka generacja PV powodująca przepływ zwrotny przez transformator do sieci SN.

PV nie ma być ograniczane jako osobna strategia sterowania. Dopuszczalne jest tylko pilnowanie ograniczenia mocy pozornej falownika:

```text
P_PV^2 + Q_PV^2 <= S_INV^2
```

Magazyn energii:

- steruje tylko mocą czynną P,
- może się ładować lub rozładowywać,
- działa osobno w fazach L1, L2, L3,
- nie steruje mocą bierną Q,
- ma stałą moc znamionową, której PSO nie może zmieniać.

Preferowana moc magazynu:

```text
P_STORAGE_TOTAL_MAX = 60 kW
P_STORAGE_PHASE_MAX = 20 kW na fazę
```

---

## 3. Przypadki badawcze

Skrypt ma obsługiwać cztery przypadki:

```text
base_no_control
pso_global
local_qu_storage_tr
local_qu_storage_end
```

### 3.1. base_no_control

Wariant bazowy bez aktywnego sterowania.

Założenia:

- PV generuje zadaną moc czynną,
- falowniki PV mają `av_mode = constq`,
- Q_PV jest brane z danych wejściowych (Excel),
- magazyn nie pracuje,
- P_storage_L1 = P_storage_L2 = P_storage_L3 = 0.

Ten wariant służy jako punkt odniesienia do liczenia poprawy procentowej.

### 3.2. pso_global

Wariant centralny, referencyjny.

PSO ma pełną informację o stanie sieci i dobiera:

- Q falowników PV,
- P magazynu w fazach L1, L2, L3.

W tym wariancie wszystkie PV mają `av_mode = constq`, a PSO ustawia Q bezpośrednio.

PSO nie może dobierać:

- mocy znamionowej magazynu,
- lokalizacji magazynu,
- curtailmentu PV,
- Q magazynu.

PSO ma korzystać z tych samych zasobów technicznych co metody lokalne. Różnica polega na tym, że PSO ma pełną informację o sieci, a metody lokalne tylko pomiary lokalne.

### 3.3. local_qu_storage_tr

Wariant lokalny z magazynem przy transformatorze.

Falowniki PV pracują według lokalnej charakterystyki Q(U) z PowerFactory (`av_mode = qvchar`).

Magazyn znajduje się przy transformatorze po stronie nn.

Magazyn korzysta z lokalnych pomiarów transformatora:

- prądy fazowe transformatora,
- moce fazowe transformatora,
- ewentualnie napięcia fazowe po stronie nn.

Uwaga: sam moduł prądu nie wystarcza do wykrycia kierunku przepływu mocy. Do wykrycia przepływu zwrotnego do sieci SN należy użyć fazowych mocy czynnych transformatora albo obliczyć moc z fazorów U i I.

Magazyn działa skokowo:

```text
0%, 25%, 50%, 75%, 100%
```

Dla 20 kW na fazę daje to:

```text
0 kW, 5 kW, 10 kW, 15 kW, 20 kW
```

Logika:

- jeśli występuje eksport mocy w danej fazie, magazyn ładuje się w tej fazie,
- im większy eksport lub większa asymetria/przeciążenie fazy, tym wyższy stopień pracy,
- wszystkie nastawy muszą być ograniczone przez Smax/Pmax magazynu.

### 3.4. local_qu_storage_end

Wariant lokalny z magazynem w głębi sieci.

Falowniki PV pracują według tej samej charakterystyki Q(U) co w wariancie przy transformatorze (`av_mode = qvchar`).

Magazyn znajduje się w węźle krytycznym, wybranym na podstawie wyników `base_no_control`.

Najprostsze kryterium wyboru węzła krytycznego:

```text
węzeł z największym napięciem fazowym Umax w wariancie bazowym
```

Magazyn mierzy lokalne napięcia fazowe i działa skokowo:

- gdy napięcie w fazie przekracza próg, magazyn ładuje się w tej fazie,
- im większe przekroczenie napięcia, tym wyższy stopień pracy,
- faza z większym lokalnym odchyleniem napięcia może dostać wyższy stopień pracy.

---

## 4. Charakterystyka Q(U) falowników PV

W wariantach lokalnych falowniki PV mają pracować według jednej wspólnej charakterystyki Q(U) z martwą strefą, ustawionej bezpośrednio w PowerFactory.

Skrypt Python nie liczy ręcznie Q(U): przełącza tylko `av_mode` PV (`constq` lub `qvchar`) zależnie od przypadku.

Proponowana logika krzywej:

```text
U <= 0.95 pu:
    Q = -Qmax

0.95 pu < U < 0.97 pu:
    Q liniowo od -Qmax do 0

0.97 pu <= U <= 1.03 pu:
    Q = 0

1.03 pu < U < 1.08 pu:
    Q liniowo od 0 do +Qmax

U >= 1.08 pu:
    Q = +Qmax
```

Konwencja znaku:

```text
Q > 0  - falownik pobiera moc bierną indukcyjną
Q < 0  - falownik oddaje moc bierną pojemnościową
```

Przy wzroście napięcia falownik powinien pobierać moc bierną, czyli dla wysokiego U powinno być Q > 0.

Dla każdego falownika i fazy należy pilnować ograniczenia:

```text
Qmax_available = sqrt(S_INV^2 - P_PV^2)
Q_command = clamp(Q_from_QU_curve, -Qmax_available, Qmax_available)
```

---

## 5. Co ma robić skrypt

Skrypt powinien umożliwiać:

1. uruchomienie wariantu bazowego,
2. uruchomienie wariantu PSO,
3. uruchomienie lokalnej regulacji Q(U),
4. uruchomienie skokowej reguły magazynu przy transformatorze,
5. uruchomienie skokowej reguły magazynu w głębi sieci,
6. eksport wyników do CSV/XLSX,
7. obliczenie wskaźników jakości,
8. porównanie wariantów względem bazy i względem PSO.

Docelowe przypadki mają mieć osobne katalogi wynikowe:

```text
results/base_no_control/
results/pso_global/
results/local_qu_storage_tr/
results/local_qu_storage_end/
```

---

## 6. Wskaźniki wynikowe

Dla każdego wariantu należy policzyć co najmniej:

- Umax_pu,
- Umin_pu,
- średnie odchylenie napięcia od 1,0 pu,
- maksymalne odchylenie napięcia od 1,0 pu,
- maksymalną różnicę napięć między fazami w węźle,
- wskaźnik asymetrii napięć,
- prądy fazowe transformatora,
- asymetrię prądów transformatora,
- moc czynną transformatora w fazach,
- eksport mocy czynnej do sieci SN,
- straty mocy czynnej,
- sumaryczne wykorzystanie Q przez PV,
- moc magazynu w fazach L1, L2, L3,
- informację, czy przekroczono ograniczenia Smax/Pmax.

Porównanie względem bazy:

```text
improvement_vs_base = (X_base - X_case) / X_base * 100%
```

Skuteczność względem PSO:

```text
effectiveness_vs_pso = (X_base - X_local) / (X_base - X_pso) * 100%
```

Stosować tylko dla wskaźników, które mają być minimalizowane, np. asymetria, eksport, odchylenie napięcia, straty.

---

## 7. Najważniejsze zasady kontroli jakości

Agent sprawdzający skrypt ma szczególnie zweryfikować:

1. Czy wszystkie obliczenia są fazowe/asymetryczne.
2. Czy skrypt poprawnie pilnuje Smax falowników PV.
3. Czy skrypt poprawnie pilnuje Pmax/Smax magazynu.
4. Czy PSO używa tych samych zasobów technicznych co metoda lokalna.
5. Czy PSO nie ma dodatkowych zmiennych niedostępnych lokalnie.
6. Czy kierunek przepływu przez transformator jest rozpoznawany na podstawie mocy, a nie samego modułu prądu.
7. Czy warianty lokalne działają skokowo, a nie ciągle.
8. Czy moc magazynu jest taka sama w wariantach lokalnych i PSO.
9. Czy eksport wyników pozwala porównać wszystkie przypadki.
10. Czy wyniki można odtworzyć z konfiguracji bez ręcznego poprawiania kodu.

Najważniejsza zasada metodologiczna:

**PSO może mieć przewagę informacyjną, ale nie może mieć innych zasobów technicznych niż metody lokalne.**

---

## 8. Oczekiwane zadanie dla kolejnego agenta

Kolejny agent ma:

1. przejrzeć istniejący skrypt,
2. sprawdzić, które elementy już działają,
3. wskazać, które funkcje są brakujące lub błędne,
4. poprawić strukturę przypadków badawczych,
5. dopilnować ograniczeń Smax/Pmax,
6. sprawdzić logikę Q(U),
7. sprawdzić logikę skokowego magazynu,
8. sprawdzić zmienne decyzyjne PSO,
9. przygotować instrukcję uruchamiania,
10. zaproponować kolejne kroki implementacyjne.

Nie należy rozszerzać zakresu badań bez potrzeby. Priorytetem jest spójne i uczciwe porównanie czterech przypadków:

```text
base_no_control
pso_global
local_qu_storage_tr
local_qu_storage_end
```
