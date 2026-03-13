def get_coverage(art, qta, stocks):
    """
    Restituisce il codice dell'articolo che copre il fabbisogno.
    Se copre il padre, restituisce il padre. 
    Se copre il figlio, restituisce il figlio.
    """
    if art not in stocks: 
        return None
    
    # 1. Prova con la giacenza del Padre
    if stocks[art]['GIA'] >= qta:
        stocks[art]['GIA'] -= qta
        return art # Restituisce una stringa
    
    # 2. Se non basta, prova con il Figlio
    figlio = stocks[art].get('FIGLIO', 'NAN')
    if figlio != 'NAN' and figlio in stocks:
        if stocks[figlio]['GIA'] >= qta:
            stocks[figlio]['GIA'] -= qta
            return figlio # Restituisce una stringa
            
    return None
