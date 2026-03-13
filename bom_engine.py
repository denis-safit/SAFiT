def get_coverage(art, qta, stocks):
    if art not in stocks: 
        return None
    
    # 1. Prova col Padre
    if stocks[art]['GIA'] >= qta:
        stocks[art]['GIA'] -= qta
        return str(art).strip().upper() # Restituisce solo la stringa pulita
    
    # 2. Prova col Figlio
    figlio = str(stocks[art].get('FIGLIO', 'NAN')).strip().upper()
    if figlio != 'NAN' and figlio in stocks:
        if stocks[figlio]['GIA'] >= qta:
            stocks[figlio]['GIA'] -= qta
            return figlio # Restituisce solo la stringa pulita
            
    return None
