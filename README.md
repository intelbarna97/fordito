# Táblázatos elemző gráf kirajzolással
## 1. Működése

- **Táblázat**
  A programmal be tudunk olvasni egy táblázatot ami a nyelvtan szabályait tartalmazzák.
  Ennek a táblázatnak az elérési útját kell bemásolni a megfelelő inputra.

  A program a következő nyelvtan alapján van megírva:
  '''
  G = ({E, E’, T, T’, F}, {+, *, (, ), i}, P, E)

  S -> E#
  E -> TE’
  E’ -> +TE’ | e
  T -> FT’
  T’ -> *FT’ | e
  F -> (E) | i
  '''
  Ezt a táblázatot módosítani és menteni(felülírja az eredetit) is lehet a program segítségével.
- **Input**
  Szükség van egy bemenetre, amit a program átalakít a megfelelő alakra.
  Pl.: 
    bemenet:                ( 3 * 3 ) + 2
    átalakított bemenet:    ( i + i ) * i 

- **Epszilon**
  Van egy üres string, ezt meg lehet adni az epszilon mezőben. 
  Alapértelmezetten: e

- **Lista**
A lista a program lépéseit tartalmazza egy rendezett hármasban.
 - az első elem az aktuális input szalag maradék része
 - a középső elem a verem aktuális tartalma
 - a jobboldali elem pedig az eddig alkalmazott szabályok sorozata

- **Szintaxisfa**
  Ha az algoritmus végigért akkor a program felépít egy fát a lépések alapján, majd ezt kirajzolja a canvasre.
  Zöld színnel jeleníti meg azokat a leveleket amik az input karakterei.
  
  
## 2. Algoritmus

  - Átalakítom az inputot a megfelelő alakra.
  - A beolvasott táblázat elemeiből az inputról beolvasottak alapján kiválasztom a megfelelő szabályt(x,y koordinátákat használva).
  - A kiolvasott szabály alapján kiszámolom a lépést és hozzáadom a listához, addig ismételve amíg nem áll meg a szabályok alapján.
  - Miután megállt az elemzés a program felépíti a szintaxisfát a lépések alapján, majd ezt kirajzolja.
