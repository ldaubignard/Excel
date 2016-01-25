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
