"""Portable calculation/preview companion. Does not execute CorelDRAW VBA."""
import argparse
import json
import math
from pathlib import Path

ROOT = Path(__file__).parent


def plan(width=50., height=40., web=220., rapport=300., distortion=.97,
         columns=None, rows=None, reverse=False, clearance=1.5, rail=2.):
    values=(width,height,web,rapport,distortion,clearance,rail)
    if not all(isinstance(v,(int,float)) and math.isfinite(v) for v in values):
        raise ValueError('All dimensions must be finite numbers')
    if min(width,height,web,rapport,rail)<=0 or clearance<0 or not 0<distortion<=1:
        raise ValueError('Positive dimensions and compression 0 < k <= 1 required')
    max_columns=math.floor((web-2*(clearance+rail)+1e-8)/width)
    max_rows=math.floor((rapport+1e-8)/height)
    if max_columns<1 or max_rows<1:
        raise ValueError('Label does not fit the selected web/rapport')
    columns=max_columns if columns is None else columns
    rows=max_rows if rows is None else rows
    for count,maximum in ((columns,max_columns),(rows,max_rows)):
        if not isinstance(count,int) or isinstance(count,bool) or not 1<=count<=maximum:
            raise ValueError('Rows/columns must be integers within printable limits')
    if columns*rows>10000:
        raise ValueError('Demonstration is limited to 10000 labels')
    left=(web-columns*width)/2
    pitch=rapport/rows
    labels=[{'column':c+1,'row':r+1,'x':left+c*width,'y':r*pitch*distortion,
             'width':width,'height':height*distortion}
            for r in range(rows) for c in range(columns)]
    rails=[{'x':left-clearance-rail,'width':rail},
           {'x':left+columns*width+clearance,'width':rail}]
    return {'columns':columns,'rows':rows,'count':columns*rows,'web':web,'rapport':rapport,
            'height':rapport*distortion,'distortion':distortion,'reverse':reverse,
            'max_columns':max_columns,'max_rows':max_rows,'clearance':clearance,
            'labels':labels,'rails':rails,'pitch_before_distortion':pitch}


def svg(p):
    # Source art is generated, not supplied by a customer or derived from any print file.
    out=[f'<svg xmlns="http://www.w3.org/2000/svg" width="{p["web"]}mm" height="{p["height"]}mm" viewBox="0 0 {p["web"]} {p["height"]}">',
         '<rect width="100%" height="100%" fill="#fbfbf6"/>']
    if p['reverse']: out.append(f'<g transform="translate({p["web"]} 0) scale(-1 1)">')
    for rail in p['rails']:
        out.append(f'<rect x="{rail["x"]}" y="0" width="{rail["width"]}" height="{p["height"]}" fill="#263d48"/>')
    for label in p['labels']:
        x,y,w,h=(label[k] for k in ('x','y','width','height'))
        out.extend([f'<g transform="translate({x} {y})">',
                    f'<rect x="0.2" y="0.2" width="{w-.4}" height="{h-.4}" rx="2" fill="#e4eddf" stroke="#70927a" stroke-width=".25"/>',
                    f'<path d="M 4 8 L 10 4 L 16 8 L 10 12 Z" fill="#48785e"/>',
                    f'<text x="4" y="{h*.55}" font-family="Arial,sans-serif" font-size="4" fill="#254b37">LABEL LAB</text>',
                    f'<text x="4" y="{h*.76}" font-family="Arial,sans-serif" font-size="2.5" fill="#4c6655">{label["column"]:02} / {label["row"]:02} · DEMO</text></g>'])
    if p['reverse']: out.append('</g>')
    out.append('</svg>')
    return '\n'.join(out)


def preview_page(p):
    return f'''<!doctype html><html lang="ru"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Label Imposition Lab</title>
<style>body{{margin:0;background:#eff2ed;color:#244236;font:15px system-ui}}main{{max-width:1120px;margin:32px auto;padding:0 24px}}h1{{font-size:32px;letter-spacing:-1px}}.grid{{display:grid;grid-template-columns:1fr 360px;gap:36px}}section{{background:white;padding:24px;border-radius:14px}}svg{{height:680px;max-width:100%;display:block;margin:auto}}.badge{{padding:10px 14px;background:#dce9d8;border-radius:22px;display:inline-block}}dt{{color:#657a6e;margin-top:22px}}dd{{margin:5px 0;font-size:22px}}.notice{{padding:16px;background:#fff4db;line-height:1.7;border-radius:10px;margin:20px 0}}small{{letter-spacing:2px}}@media(max-width:800px){{.grid{{display:block}}section{{margin-bottom:20px}}svg{{height:auto;width:100%}}}}</style>
<main><small>PRINT AUTOMATION / 02</small><h1>Раскладка этикеток по рапорту</h1><p>Расчётная демонстрация и исходный VBA-модуль для CorelDRAW</p><div class="notice">Превью сгенерировано Python. Это не скриншот запуска CorelDRAW и не готовый файл для печати. Машина DEMO-220 и коэффициент 0,97 вымышлены.</div>
<div class="grid"><section>{svg(p)}</section><section><span class="badge">{p['count']} этикеток · расчёт проверен</span><dl><dt>Раскладка</dt><dd>{p['columns']} ручья × {p['rows']} этикеток</dd><dt>Ширина полотна</dt><dd>{p['web']:g} мм</dd><dt>Рапорт до сжатия</dt><dd>{p['rapport']:g} мм</dd><dt>После дисторции</dt><dd>{p['height']:g} мм</dd><dt>Коэффициент по длине</dt><dd>{p['distortion']:g}</dd><dt>Зазор от этикетки до рельса</dt><dd>{p['clearance']:g} мм</dd><dt>Печать</dt><dd>{'Обратная · зеркало X' if p['reverse'] else 'Прямая'}</dd></dl><p>Сетка, рельсы и оформление — демонстрационные. Для production нужны параметры машин и контрольный макет типографии.</p></section></div></main></html>'''


if __name__=='__main__':
    ap=argparse.ArgumentParser()
    ap.add_argument('--reverse',action='store_true')
    ap.add_argument('--out',default='examples')
    args=ap.parse_args()
    p=plan(reverse=args.reverse)
    target=Path(args.out); target.mkdir(parents=True,exist_ok=True)
    (target/'layout.json').write_text(json.dumps(p,indent=2,ensure_ascii=False),encoding='utf-8')
    (target/'layout.svg').write_text(svg(p),encoding='utf-8')
    (target/'preview.html').write_text(preview_page(p),encoding='utf-8')
    print(f'{p["columns"]} x {p["rows"]} = {p["count"]} labels; {p["height"]} mm after distortion')
