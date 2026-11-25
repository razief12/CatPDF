from pathlib import Path
import subprocess, zipfile, io, html, re, warnings
import fitz
from PIL import Image, ImageOps
from docx import Document
from docx.shared import Pt, Inches
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches as PI
from weasyprint import HTML

KINDS={'pdfdocx':'docx','pdfxlsx':'xlsx','pdfpptx':'pptx','pdfimg':'zip','pdftxt':'txt','pdfhtml':'html','pdfmd':'md','docxpdf':'pdf','xlsxpdf':'pdf','pptxpdf':'pdf','imgpdf':'pdf','txtpdf':'pdf','htmlpdf':'pdf','jpgpng':'png','pngjpg':'jpg','webp':'webp'}
Image.MAX_IMAGE_PIXELS=20_000_000
warnings.simplefilter("error", Image.DecompressionBombWarning)

def no_fetch(url,*args,**kwargs):
    raise ValueError('External resources disabled')

def prepare_excel_layout(src):
    # Modify only the temporary upload, never the user's original workbook.
    wb = load_workbook(src)
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != 'visible':
                continue
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.scale = None
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            if not ws.page_setup.paperSize:
                ws.page_setup.paperSize = ws.PAPERSIZE_A4
            if ws.max_column >= 6:
                ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            elif not ws.page_setup.orientation:
                ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.col_breaks.brk = []
        wb.save(src)
    finally:
        wb.close()

def office(src,out):
    with zipfile.ZipFile(src) as z:
        if len(z.infolist())>5000 or sum(i.file_size for i in z.infolist())>150_000_000: raise ValueError('Office archive too large')
        if any('vbaProject' in i.filename or 'externalLinks/' in i.filename for i in z.infolist()): raise ValueError('Macros and external workbook links are not supported')
    if src.suffix.lower() == '.xlsx': prepare_excel_layout(src)
    p=subprocess.run(['libreoffice','-env:UserInstallation='+ (out.parent/'profile').as_uri(),'--headless','--convert-to','pdf','--outdir',str(out.parent),str(src)],capture_output=True,timeout=90)
    generated=src.with_suffix('.pdf')
    if p.returncode or not generated.exists(): raise ValueError('Office rendering failed')
    generated.replace(out)

def page_blocks(page):
    tables=page.find_tables(strategy="lines_strict").tables or page.find_tables().tables
    blocks=[]
    for t in tables: blocks.append((t.bbox[1],t.bbox[0],'table',t))
    for b in page.get_text('blocks',sort=True):
        if b[6]!=0: continue
        # Keep text outside detected table bounds, including partially overlapping blocks.
        lines=[]
        for line in page.get_text('dict',clip=fitz.Rect(b[:4]))['blocks']:
            for l in line.get('lines',[]):
                r=fitz.Rect(l['bbox']); center=( (r.x0+r.x1)/2,(r.y0+r.y1)/2 )
                if not any(fitz.Rect(t.bbox).contains(fitz.Point(*center)) for t in tables): lines.append(''.join(s['text'] for s in l['spans']))
        if lines: blocks.append((b[1],b[0],'text','\n'.join(lines)))
    return sorted(blocks,key=lambda b:(b[0],b[1]))

def clean_cell(v):
    return re.sub(r'\b(SECTIO)\s*\n\s*N\b', 'SECTION', re.sub(r'\bSUBTO\s*\n\s*TAL\b','SUBTOTAL',v))

def cell_style(page,rect,drawings,allspans):
    if not rect: return None,'000000',False
    r=fitz.Rect(rect); pt=fitz.Point((r.x0+r.x1)/2,(r.y0+r.y1)/2); bg=None
    for d in drawings:
        if d.get('fill') and d['rect'].contains(pt):
            bg=''.join(f'{round(c*255):02X}' for c in d['fill'][:3])
    spans=[s for s in allspans if r.contains(fitz.Point((s['bbox'][0]+s['bbox'][2])/2,(s['bbox'][1]+s['bbox'][3])/2))]
    s=max(spans,key=lambda x:len(x['text']),default={})
    return bg,f"{s.get('color',0):06X}",bool(s.get('flags',0)&16)

def legacy_html(soup,pages):
    parts=[]
    for i,p in enumerate(pages):
        w,h=re.search(r'width:([\d.]+)pt;height:([\d.]+)pt',p['style']).groups()
        p['id']=f'page{i}'; p['style']+=f';position:relative;page:cat{i};break-after:page'
        for line in p.find_all('p',recursive=False):
            line['style']=line.get('style','')+';position:absolute;margin:0;white-space:nowrap'
            for span in line.find_all('span'):
                span['style']=span.get('style','').replace('color:#ffffff','color:#000000')
        parts.append(f'<style>@page cat{i}{{size:{w}pt {h}pt;margin:0}}</style>'+str(p))
    return '<!doctype html><meta charset=utf-8><style>body{margin:0}div:last-child{break-after:auto}</style>'+''.join(parts)

def unique_svg(svg, page):
    prefix=f'p{page}_'
    svg=re.sub(r'\bid="([^"]+)"',lambda m:f'id="{prefix}{m[1]}"',svg)
    svg=re.sub(r'(href="\#|url\(#)([^"\)]+)',lambda m:m[1]+prefix+m[2],svg)
    return svg

def convert(kind,src,out):
    if kind in ('docxpdf','xlsxpdf','pptxpdf'): office(src,out); return
    if kind in ('txtpdf','htmlpdf'):
        text=src.read_text(encoding='utf-8-sig');
        if kind=='txtpdf': text='<meta charset=utf-8><pre style="white-space:pre-wrap;overflow-wrap:anywhere">'+html.escape(text)+'</pre>'
        if kind=='htmlpdf':
            soup=BeautifulSoup(text,'html.parser')
            # Repair legacy PyMuPDF exports only, not arbitrary HTML.
            pages=soup.find_all('div',id=re.compile(r'^page\d+$'),recursive=False)
            if pages and all(re.search(r'width:[\d.]+pt;height:[\d.]+pt',x.get('style','')) for x in pages):
                text=legacy_html(soup,pages)
            else:
                text='<style>@page{size:A4;margin:12mm}img,svg{max-width:100%}table{max-width:100%;overflow-wrap:anywhere}pre{white-space:pre-wrap;overflow-wrap:anywhere}</style>'+text
        HTML(string=text,url_fetcher=no_fetch).write_pdf(out); return
    if kind in ('jpgpng','pngjpg','webp','imgpdf'):
        with Image.open(src) as im:
            im=ImageOps.exif_transpose(im); im.load()
            if kind in ('pngjpg','imgpdf'):
                rgba=im.convert('RGBA'); bg=Image.new('RGB',rgba.size,'white'); bg.paste(rgba,mask=rgba.getchannel('A')); im=bg
            else: im=im.convert('RGBA' if 'A' in im.getbands() or 'transparency' in im.info else 'RGB')
            im.save(out,format={'jpgpng':'PNG','pngjpg':'JPEG','webp':'WEBP','imgpdf':'PDF'}[kind],**({'quality':95} if kind=='pngjpg' else {'lossless':True} if kind=='webp' else {})); return
    with fitz.open(src) as pdf:
        if pdf.needs_pass or not 1<=len(pdf)<=100: raise ValueError('Encrypted PDF or page limit exceeded (100)')
        if kind=='pdftxt':
            pages=[p.get_text('text',sort=True) for p in pdf]
            if any(not t.strip() and p.get_images() for p,t in zip(pdf,pages)):
                raise ValueError('Image-only page detected; OCR required')
            out.write_text('\n\f\n'.join(pages),encoding='utf-8'); return
        if kind=='pdfimg':
            with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
                for i,p in enumerate(pdf): z.writestr(f'page-{i+1:03}.png',p.get_pixmap(matrix=fitz.Matrix(min(2, 2200/max(p.rect.width,p.rect.height)),min(2, 2200/max(p.rect.width,p.rect.height))),alpha=False).tobytes('png'))
            return
        if kind=='pdfpptx':
            deck=Presentation(); r=pdf[0].rect; deck.slide_width=PI(10); deck.slide_height=PI(10*r.height/r.width)
            for p in pdf:
                slide=deck.slides.add_slide(deck.slide_layouts[6]); r=p.rect; k=min(deck.slide_width/r.width,deck.slide_height/r.height); w=int(r.width*k); h=int(r.height*k); slide.shapes.add_picture(io.BytesIO(p.get_pixmap(matrix=fitz.Matrix(min(2, 2200/max(p.rect.width,p.rect.height)),min(2, 2200/max(p.rect.width,p.rect.height))),alpha=False).tobytes('png')),int((deck.slide_width-w)/2),int((deck.slide_height-h)/2),width=w,height=h)
            deck.save(out); return
        if any(not p.get_text().strip() and p.get_images() for p in pdf): raise ValueError('Image-only page detected; OCR required to avoid silent data loss')
        if not any(p.get_text().strip() for p in pdf): raise ValueError('No selectable text; OCR required')
        if kind=='pdfhtml':
            parts=[]
            for i,p in enumerate(pdf):
                # SVG preserves vector artwork, backgrounds and text positioning.
                w,h=p.rect.width,p.rect.height
                parts.append(f'<style>@page cat{i}{{size:{w}pt {h}pt;margin:0}}</style><section style="page:cat{i};width:{w}pt;height:{h}pt;break-after:page">'+unique_svg(p.get_svg_image(text_as_path=False),i)+'</section>')
            out.write_text('<!doctype html><html><head><meta charset="utf-8"><style>body{margin:0}svg{display:block;width:100%;height:100%}section:last-child{break-after:auto}</style></head><body>'+''.join(parts)+'</body></html>',encoding='utf-8'); return
        doc=Document() if kind=='pdfdocx' else None; wb=Workbook() if kind=='pdfxlsx' else None; md=[]
        if wb: wb.remove(wb.active)
        for i,p in enumerate(pdf):
            blocks=page_blocks(p)
            drawings=p.get_drawings(); allspans=[s for b in p.get_text('dict')['blocks'] for l in b.get('lines',[]) for s in l['spans'] if s['text'].strip()]
            if doc and i: doc.add_page_break()
            if doc and i==0:
                sec=doc.sections[0]; sec.page_width=Pt(p.rect.width); sec.page_height=Pt(p.rect.height); sec.left_margin=sec.right_margin=Pt(24); sec.top_margin=sec.bottom_margin=Pt(24)
            if wb: ws=wb.create_sheet(f'Page {i+1}'); row=1
            md.append(f'## Page {i+1}')
            for _,__,typ,data in blocks:
                if typ=='text':
                    if doc: doc.add_paragraph(data)
                    if wb: ws.cell(row,1,data).data_type='s'; row+=1
                    md.append(data); continue
                t=data; data=t.extract(); n=t.col_count
                if not n: continue
                widths=[max(1,min((rr.cells[ci][2]-rr.cells[ci][0] for rr in t.rows if rr.cells[ci] is not None),default=(t.bbox[2]-t.bbox[0])/n)) for ci in range(n)]
                if doc:
                    table=doc.add_table(rows=len(data),cols=n); table.style='Table Grid'; table.autofit=False
                    scale=min(1,(p.rect.width-48)/sum(widths))
                    for col,w in zip(table.columns,widths): col.width=Pt(w*scale)
                    for rr in table.rows:
                        for cell,w in zip(rr.cells,widths): cell.width=Pt(w*scale)
                if wb:
                    for ci,w in enumerate(widths):
                        from openpyxl.utils import get_column_letter
                        ws.column_dimensions[get_column_letter(ci+1)].width=max(5,w/5.5)
                for ri,values in enumerate(data):
                    for ci,v in enumerate(values):
                        v=clean_cell(v or '')
                        bg,fg,bold=cell_style(p,t.rows[ri].cells[ci],drawings,allspans)
                        if doc:
                            cell=table.cell(ri,ci); cell.text=v
                            for para in cell.paragraphs:
                                para.paragraph_format.space_after=Pt(0)
                                for run in para.runs:
                                    run.font.size=Pt(8); run.bold=bold; run.font.color.rgb=RGBColor.from_string(fg)
                            if bg:
                                sh=OxmlElement('w:shd'); sh.set(qn('w:fill'),bg); cell._tc.get_or_add_tcPr().append(sh)
                        if wb:
                            c=ws.cell(row+ri,ci+1,v); c.data_type='s'; c.alignment=Alignment(wrap_text=True,vertical='top'); c.border=Border(**{s:Side(style='thin',color='BBBBBB') for s in ('left','right','top','bottom')}); c.font=Font(bold=bold,color=fg); c.fill=PatternFill('solid',fgColor=bg) if bg else PatternFill();
                            if ci>0 and re.fullmatch(r'RM\s*-?[\d,]+(?:\.\d+)?',v):
                                c.value=float(v[2:].strip().replace(',','')); c.number_format='"RM" #,##0.00'
                            if ci>0 and re.fullmatch(r'-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?',v) and len(v.replace(',','').replace('.','').replace('-',''))<=15:
                                header=(data[0][ci] or '').lower();
                                if re.search(r'qty|quantity|rate|amount|times|weight|factor|\(m\)',header):
                                    c.value=float(v.replace(',','')) if '.' in v else int(v.replace(',','')); c.number_format='#,##0.'+'0'*len(v.rsplit('.',1)[1]) if '.' in v else '#,##0'
                    vals=[clean_cell(v or '').replace('|','\\|').replace('\n','<br>') for v in values]; md.append('| '+' | '.join(vals+['']*(n-len(vals)))+' |')
                    if ri==0: md.append('| '+' | '.join(['---']*n)+' |')
                # Merge only cells whose detected geometry spans adjacent columns.
                for ri,rr in enumerate(t.rows):
                    for ci,rect in enumerate(rr.cells):
                        if rect is None: continue
                        end=ci
                        while end+1<n and rr.cells[end+1] is None and rect[2]>=t.bbox[0]+sum(widths[:end+2])-2: end+=1
                        if end>ci:
                            if doc: table.cell(ri,ci).merge(table.cell(ri,end))
                            if wb: ws.merge_cells(start_row=row+ri,start_column=ci+1,end_row=row+ri,end_column=end+1)
                if wb: row+=len(data)+1
                md.append('')
            if wb:
                for col in ws.columns:
                    pass # table geometry determines widths
        if doc: doc.save(out)
        elif wb: wb.save(out)
        else: out.write_text('\n\n'.join(md).replace(' |\n\n|',' |\n|'),encoding='utf-8')
