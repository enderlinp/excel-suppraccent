# Excel - SUPPRACCENT()
Fonction personnalisée Excel permettant de supprimer les accents d'une chaîne de texte passée en paramètre.

## Présentation
Visual Basic pour Applications (VBA), le langage de programmation intégré à toute la suite Microsoft Office, est une version édulcorée de [Visual Basic](https://fr.wikipedia.org/wiki/Visual_Basic). C'est ce langage que nous allons utiliser aujourd'hui pour créer une **fonction personnalisée** absente d'Excel et pourtant bien utile.

Cette fonction, que j'ai intitulée `SUPPRACCENT()`, permet de supprimer tous les caractères accentués d'une chaîne de texte (passée en paramètre) par leurs équivalents sans accents.

Son comportement s'apparente à celui de la fonction [strtr()](https://www.php.net/manual/fr/function.strtr.php) en PHP.

A partir d'Excel, ouvrez **Visual Basic Editor** (ALT+F11 sur PC et Mac) et insérez un nouveau module à partir de l'explorateur de projets. Dans l'éditeur, copiez/collez les lignes ci-dessous, enregistrez le module puis fermez l'éditeur.

```bas
' Déclaration obligatoire des variables
Option Explicit

' -------------------------------------------------------------------------------------------------------------
' Fonction personnalisée permettant de supprimer les accents au sein d’une chaine de texte passée en paramètre.
'
' Auteur : enderlinp
' -------------------------------------------------------------------------------------------------------------

Function SUPPRACCENT(str As String) As String

    ' Déclaration des variables
    Dim strAccent, strNoAccent, strFrom, strTo As String, i As Integer
    
    ' Liste des caractères accentués et leurs équivalents non accentués
    strAccent = "àâäçéèêëîïôöùûüÿÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸ"
    strNoAccent = "aaaceeeeiioouuuyAAACEEEEIIOOUUUY"
    
    ' On récupère un caractère de "strAccent" et de "strNoAccent" à la position i à chaque itération
    For i = 1 To Len(strAccent)
        
        strFrom = Mid(strAccent, i, 1)
        strTo = Mid(strNoAccent, i, 1)
        
        ' On remplace dans la chaîne "str" les caractères accentués par leurs équivalents non accentués
        str = Replace(str, strFrom, strTo)
        
    Next
    
    ' On retourne la chaîne "str" sans accents
    SUPPRACCENT = str

End Function
```

De retour dans Excel, la nouvelle fonction est immédiatement disponible, comme n'importe quelle fonction intégrée d'Excel.
