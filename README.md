# Excel
Mes fonctions Excel

### Récupérer une adresse email entre <>
Exemple de chaîne de caractère dans la cellule A2
```sh
Louis DAUBIGNARD <louis.daubignard@xxxxx.com>
```
Formule Excel :
```sh
=SI(ESTERREUR(CHERCHE("<";A2))=FAUX;STXT(A2;TROUVE("<";A2)+1;TROUVE(">";A2;TROUVE("<";A2))-TROUVE("<";A2)-1);"")
```
ou si besoin de récupérer les adresses email sans <> 
```sh
=SI(ESTERREUR(CHERCHE("<";A2))=FAUX;STXT(A2;TROUVE("<";A2)+1;TROUVE(">";A2;TROUVE("<";A2))-TROUVE("<";A2)-1);SI(ESTERREUR(CHERCHE("@";A2))=FAUX;A2;""))
```

Résultat :
```sh
louis.daubignard@xxxxx.com
```

### Récupérer la chaîne de caractère avant le @ d'une adresse email 
Exemple de chaîne de caractère dans la cellule C2
```sh
louis.daubignard@xxxxx.com
```
Formule Excel :
```sh
=GAUCHE(C2;TROUVE("@";C2)-1)
```

Résultat :
```sh
louis.daubignard
```

### Récupérer le ndd après le @ d'une adresse email 
Exemple de chaîne de caractère dans la cellule C2
```sh
louis.daubignard@xxxxx.com
```
Formule Excel :
```sh
=STXT(C2;[@depart];[@[longueur texte]])
    [@depart] =TROUVE("@";C2)+1
    [@[longueur texte]] =NBCAR(C2)-[@depart]+1
```
```sh
=STXT(C2;TROUVE("@";C2)+1;NBCAR(C2)-(TROUVE("@";C2)+1)+1)
```

Résultat :
```sh
xxxxx.com
```

### Récupérer un prénom d'une adresse email 
Exemple de chaîne de caractère dans la cellule C2
```sh
louis.daubignard@xxxxx.com
```
Formule Excel :
```sh
=NOMPROPRE(GAUCHE(GAUCHE(C2;TROUVE("@";C2)-1);TROUVE(".";GAUCHE(C2;TROUVE("@";C2)-1))-1))
```

Résultat :
```sh
Louis
```

### Récupérer un nom d'une adresse email 
Exemple de chaîne de caractère dans la cellule C2
```sh
louis.daubignard@xxxxx.com
```
Formule Excel :
```sh
=MAJUSCULE(STXT(GAUCHE(C2;TROUVE("@";C2)-1);TROUVE(".";GAUCHE(C2;TROUVE("@";C2)-1))+1;(NBCAR(GAUCHE(C2;TROUVE("@";C2)-1))-TROUVE(".";GAUCHE(C2;TROUVE("@";C2)-1)))))
```

Résultat :
```sh
DAUBIGNARD
```

### Récupérer un code entre crochet et générer une url avec Id 
Exemple de chaîne de caractère dans la cellule D5
```sh
[7591] ceci est une chaîne de caractère
```
Formule Excel :
```sh
=SI(ESTERREUR(CHERCHE("[";D5))=FAUX;"https://client.zendesk.com/agent/tickets/"&STXT(D5;2;4);"")
```
Résultat :
```sh
https://client.zendesk.com/agent/tickets/7591
```

### Récupérer une chaîne de caractère avant un caractère (-)
Exemple de chaîne de caractère dans la cellule D5
```sh
VCV10L-AEPGKW  
```
Formule Excel :
```sh
=GAUCHE(D5;TROUVE("-";D5)-1)
```
Résultat :
```sh
VCV10L
```

### Vérifier si une valeur est numérique
Exemple de chaîne de caractère dans la cellule D5
```sh
VCV10L-AEPGKW 
```
Formule Excel :
```sh
=SI(ESTNUM(CHERCHE("-";D5));GAUCHE(D5;1)&STXT(D5;CHERCHE("-";D5)+1;1);GAUCHE(D5;1))
```
Résultat :
```sh
VA
```

### Formatter une chaîne de carctère en format monétaire
Exemple de chaîne de caractère dans la cellule D5
```sh
ANNEE FISCALE 2017 : 12426
```
Formule Excel :
```sh
="ANNEE FISCALE 2017" &" : "&TEXTE(SOMME(F11:N11;F13:N13;F15:N15);"# ##0 €")
```
Résultat :
```sh
ANNEE FISCALE 2017 : 12 426 €
```

Formule | Description
:-- | --: 
=TEXTE(1234,567;"# ##0,00 €")     | Devise avec un séparateur des milliers et 2 décimales : 1 234,57 €. Notez qu’Excel arrondit la valeur à 2 décimales.  
=TEXTE(AUJOURDHUI();"JJ/MM/AA")   | Date du jour au format JJ/MM/AA (par exemple, 14/03/12) 
=TEXTE(AUJOURDHUI();"JJJJ")       | Date du jour de la semaine (par exemple, Lundi)
=TEXTE(MAINTENANT();"HH:MM")      | Heure actuelle (par exemple, 13:29)
=TEXTE(0,285;"0,0 %")             | Pourcentage (par exemple, 28,5 %)
=TEXTE(4,34 ;"# ?/?")             | Fraction (par exemple, 4 1/3)
=SUPPRESPACE(TEXTE(0,34;"# ?/?")) | Fraction (par exemple, 1/3) Notez que cette formule utilise la fonction SUPPRESPACE pour supprimer l’espace de début dans le cas d’une valeur décimale.
=TEXTE(12200000;"0,00E+00")       | Notation scientifique (par exemple, 1,22E+07)  
=TEXTE(1234567898;"[<=9999999]###-####;(###) ###-####")| Spécial (numéro de téléphone) (par exemple, (123) 456-7898)
=TEXTE(1234;"0000000")            | Ajouter des zéros (0) de début (par exemple, 0001234)
=TEXTE(123456;"##0° 00' 00''")    | Personnalisée - Latitude/Longitude
