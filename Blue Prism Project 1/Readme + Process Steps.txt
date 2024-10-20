PROCES OBLICZANIA ZUŻYCIA ENERGII URZĄDZEŃ PRZY UŻYCIU KALKULATORA ENERGII 

KLUCZOWA INFORMACJA: Aby proces działał poprawnie, należy na stronie "Tworzenie kolekcji" w data item "File Name" wkleić ścieżkę, w której aktualnie znajduje się plik "Dane Urządzeń.xlsx" na urządzeniu. Najlepiej zrobić tak: kliknąć na ikonę danego pliku, a następnie wcisnąć "Ctrl+Shift+C" - ścieżka zostanie automatycznie skopiowana. 
1. Otwarcie Excela
2. Otwarcie pliku Excela "Dane Urządzeń.xlsx"
3. Skopiowanie danych z pliku "Dane Urządzeń.xlsx" z arkusza "Arkusz1"
4. Wklejenie skopiowanych danych do kolekcji "Dane" na Main Page
5. Włączenie przeglądarki Chrome 
6. Otworzenie strony z kalkulatorem energii - https://www.tauron-dystrybucja.pl/kalkulator-zuzycia-pradu
7. Usunięcie danych znajdujących się domyślnie w polu "Moc urządzenia"
8. Usunięcie danych znajdujących się domyślnie w polu "Czas używania urządzenia"
9. Wpisanie danych "Moc urządzenia" z kolekcji "Dane" w pole "Moc urządzenia" dla urządzenia 1
10. Wpisanie danych "Czas używania urządzenia" z kolekcji "Dane" w pole "Czas używania urządzenia" dla urządzenia 1
11. Obliczenie zużycia prądu i jego kosztu
12. Skopiowanie obliczonej wartości zużycia prądu do procesu w data item "Urządzenie 1 - Zużycie" oraz do kolumny "Zużycie" w kolekcji "Dane"
13. Skopiowanie obliczonej wartości kosztu zużycia prądu do procesu w data item "Urządzenie 1 - Koszt" oraz do kolumny "Koszt" w kolekcji "Dane".
14. Usunięcie danych znajdujących się w polu "Moc urządzenia"
15. Usunięcie danych znajdujących się w polu "Czas używania urządzenia"
16. Wpisanie danych "Moc urządzenia" z kolekcji "Dane" w pole "Moc urządzenia" dla urządzenia 2
17. Wpisanie danych "Czas używania urządzenia" z kolekcji "Dane" w pole "Czas używania urządzenia" dla urządzenia 2
18. Obliczenie zużycia prądu i jego kosztu
19. Skopiowanie obliczonej wartości zużycia prądu do procesu w data item "Urządzenie 2 - Zużycie" oraz do kolumny "Zużycie" w kolekcji "Dane"
20. Skopiowanie obliczonej wartości kosztu zużycia prądu do procesu w data item "Urządzenie 2 - Koszt" oraz do kolumny "Koszt" w kolekcji "Dane".
21. Usunięcie danych znajdujących się w polu "Moc urządzenia"
22. Usunięcie danych znajdujących się w polu "Czas używania urządzenia"
23. Wpisanie danych "Moc urządzenia" z kolekcji "Dane" w pole "Moc urządzenia" dla urządzenia 3
24. Wpisanie danych "Czas używania urządzenia" z kolekcji "Dane" w pole "Czas używania urządzenia" dla urządzenia 3
25. Obliczenie zużycia prądu i jego kosztu
26. Skopiowanie obliczonej wartości zużycia prądu do procesu w data item "Urządzenie 3 - Zużycie" oraz do kolumny "Zużycie" w kolekcji "Dane"
27. Skopiowanie obliczonej wartości kosztu zużycia prądu do procesu w data item "Urządzenie 3 - Koszt" oraz do kolumny "Koszt" w kolekcji "Dane".
28. Usunięcie danych znajdujących się w polu "Moc urządzenia"
29. Usunięcie danych znajdujących się w polu "Czas używania urządzenia"
30. Wpisanie danych "Moc urządzenia" z kolekcji "Dane" w pole "Moc urządzenia" dla urządzenia 4
31. Wpisanie danych "Czas używania urządzenia" z kolekcji "Dane" w pole "Czas używania urządzenia" dla urządzenia 4
32. Obliczenie zużycia prądu i jego kosztu
33. Skopiowanie obliczonej wartości zużycia prądu do procesu w data item "Urządzenie 4 - Zużycie" oraz do kolumny "Zużycie" w kolekcji "Dane"
34. Skopiowanie obliczonej wartości kosztu zużycia prądu do procesu w data item "Urządzenie 4 - Koszt" oraz do kolumny "Koszt" w kolekcji "Dane".
35. Zamknięcie przeglądarki Chrome
36. Wklejenie danych z kolekcji "Dane" do pliku "Dane Urządzeń.xlsx", z którego wcześniej były pobierane dane, w puste kolumny "Zużycie" oraz "Koszt". 
37. Zapisanie gotowego pliku z wklejonymi danymi i zamknięcie Excela

Główny proces + kolekcja "Dane Urządzeń" znajdują się na stronie "Main Page", natomiast proces usuwania danych, wpisywania nowych oraz kopiowania ich do odpowiednich data itemów znajdujących się na "Main Page" odbywa się na stronie "Policz Zużycie Prądu". Główny proces musi podjąć 4 decyzje dotyczące numeru urządzenia, aby wpisał dane odpowiednie dla każdego z nich. W obiekcie użyte zostały dynamiczne elementy, takie jakie "web id" czy "input text". Akcje dotyczące Excela wykorzystują wbudowany w Blue Prism obiekt "MS Excel VBO". 