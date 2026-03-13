def get_coverage(art, qta, stocks, depth=0):
    if depth > 5 or art not in stocks: return None, 0
    s = stocks[art]
    if s['GIA'] >= qta:
        s['GIA'] -= qta
        return art, depth
    figlio = s.get('FIGLIO', 'NAN')
    if figlio != 'NAN' and figlio in stocks:
        return get_coverage(figlio, qta, stocks, depth + 1)
    return None, -1
