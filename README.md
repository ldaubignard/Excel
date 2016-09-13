# Excel
Mes fonctions Excel

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
