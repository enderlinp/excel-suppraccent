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
