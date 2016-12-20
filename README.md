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
