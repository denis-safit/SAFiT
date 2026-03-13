def get_coverage(art, qta, stocks, depth=0):
    """
    Motore ricorsivo: scava nella distinta base.
    depth > 5 serve a evitare loop infiniti se ci sono errori nei dati.
    """
    if depth > 5 or art not in stocks: 
        return None, 0
    
    s = stocks[art]
    
    # 1. Prova a usare la giacenza del Padre
    if s['GIA'] >= qta:
        s['GIA'] -= qta
        return art, depth
    
    # 2. Se non basta, prova a cercare nel Figlio
    figlio = s.get('FIGLIO', 'NAN')
    if figlio != 'NAN' and figlio in stocks:
        return get_coverage(figlio, qta, stocks, depth + 1)
        
    return None, -1
