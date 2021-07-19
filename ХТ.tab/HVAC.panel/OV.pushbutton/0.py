
rectsSize = []  # Размер воздуховодов
for i in rects:
    b = round(i.LookupParameter('Ширина').AsDouble() * k, 2)
    h = round(i.LookupParameter('Высота').AsDouble() * k, 2)
    if h > b:
        b, h = h, b
    if b < 251:
        s = 'δ=0,5'
    elif b < 1001:
        s = 'δ=0,7'
    else:
        s = 'δ=0,9'
    sys_name = i.LookupParameter('Имя системы').AsString()
    sys_name = sys_name if sys_name else 'Не определено'
    if 'Д' in sys_name:
        s = 'δ=1,2'
    if i.LookupParameter('Длина').AsDouble() * k > 10:
        rectsSize.append('{:.0f}×{:.0f}, {}'.format(b, h, s))
    else:
        rectsSize.append('Не учитывать {:.0f}×{:.0f}, {}'.format(b, h, s))

roundsDiameter = []  # Диаметр воздуховодов (и гибких)
for i in rounds:
    d = round(i.LookupParameter('Диаметр').AsDouble() * k, 2)
    if d < 201:
        s = 'δ=0,5'
    elif d < 451:
        s = 'δ=0,6'
    else:
        s = 'δ=0,7'
    if i.LookupParameter('Имя системы').AsString():
        if 'Д' in i.LookupParameter('Имя системы').AsString():
            s = 'δ=1,2'
    else:
        s = ''
    roundsDiameter.append('ø{:.0f}, {}'.format(d, s))


for duct in ducts:
    if duct.LookupParameter('Ширина'):
        b = round(i.LookupParameter('Ширина').AsDouble() * k, 2)
        h = round(i.LookupParameter('Высота').AsDouble() * k, 2)
        if h > b:
            b, h = h, b
        if b < 251:
            s = 'δ=0,5'
        elif b < 1001:
            s = 'δ=0,7'
        else:
            s = 'δ=0,9'
        sys_name = i.LookupParameter('Имя системы').AsString()
        if not sys_name:
            sys_name = 'Не определено'
        if 'Д' in sys_name:
            s = 'δ=1,2'
        if i.LookupParameter('Длина').AsDouble() * k > 10:
            rectsSize.append('{:.0f}×{:.0f}, {}'.format(b, h, s))
        else:
            rectsSize.append('Не учитывать {:.0f}×{:.0f}, {}'.format(b, h, s))
