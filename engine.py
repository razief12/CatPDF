from pathlib import Path
import subprocess, zipfile, io, html, re, warnings
import fitz
from PIL import Image, ImageOps
from docx import Document
from docx.shared import Pt, Inches
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from pptx import Presentation
from pptx.util import Inches as PI
from weasyprint import HTML

KINDS={'pdfdocx':'docx','pdfxlsx':'xlsx','pdfpptx':'pptx','pdfimg':'zip','pdftxt':'txt','pdfhtml':'html','pdfmd':'md','docxpdf':'pdf','xlsxpdf':'pdf','pptxpdf':'pdf','imgpdf':'pdf','txtpdf':'pdf','htmlpdf':'pdf','jpgpng':'png','pngjpg':'jpg','webp':'webp'}
Image.MAX_IMAGE_PIXELS=20_000_000
warnings.simplefilter("error", Image.DecompressionBombWarning)

def no_fetch(url,*args,**kwargs):
    raise ValueError('External resources disabled')

def office(src,out):
    with zipfile.ZipFile(src) as z:
        if len(z.infolist())>5000 or sum(i.file_size for i in z.infolist())>150_000_000: raise ValueError('Office archive too large')
        if any('vbaProject' in i.filename or 'externalLinks/' in i.filename for i in z.infolist()): raise ValueError('Macros and external workbook links are not supported')
    p=subprocess.run(['libreoffice','-env:UserInstallation='+ (out.parent/'profile').as_uri(),'--headless','--convert-to','pdf','--outdir',str(out.parent),str(src)],capture_output=True,timeout=90)
    generated=src.with_suffix('.pdf')
    if p.returncode or not generated.exists(): raise ValueError('Office rendering failed')
    generated.replace(out)

def page_blocks(page):
    tables=page.find_tables().tables
    blocks=[]
    for t in tables: blocks.append((t.bbox[1],t.bbox[0],'table',t.extract()))
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

def convert(kind,src,out):
    if kind in ('docxpdf','xlsxpdf','pptxpdf'): office(src,out); return
    if kind in ('txtpdf','htmlpdf'):
        text=src.read_text(encoding='utf-8-sig');
        if kind=='txtpdf': text='<meta charset=utf-8><pre style="white-space:pre-wrap;overflow-wrap:anywhere">'+html.escape(text)+'</pre>'
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
        if kind=='pdftxt': out.write_text('\n\f\n'.join(p.get_text(sort=True) for p in pdf),encoding='utf-8'); return
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
            out.write_text('<!doctype html><meta charset=utf-8>'+''.join(p.get_text('html') for p in pdf),encoding='utf-8'); return
        doc=Document() if kind=='pdfdocx' else None; wb=Workbook() if kind=='pdfxlsx' else None; md=[]
        if wb: wb.remove(wb.active)
        for i,p in enumerate(pdf):
            blocks=page_blocks(p)
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
                n=max(map(len,data),default=0)
                if not n: continue
                if doc: table=doc.add_table(rows=len(data),cols=n); table.style='Table Grid'
                for ri,values in enumerate(data):
                    for ci,v in enumerate(values):
                        v=v or ''
                        if doc:
                            cell=table.cell(ri,ci); cell.text=v
                            for para in cell.paragraphs:
                                para.paragraph_format.space_after=Pt(0)
                                for run in para.runs: run.font.size=Pt(8)
                        if wb:
                            c=ws.cell(row+ri,ci+1,v); c.data_type='s'; c.alignment=Alignment(wrap_text=True,vertical='top'); c.border=Border(**{s:Side(style='thin',color='BBBBBB') for s in ('left','right','top','bottom')}); c.font=Font(bold=ri==0)
                    vals=[str(v or '').replace('|','\\|').replace('\n','<br>') for v in values]; md.append('| '+' | '.join(vals+['']*(n-len(vals)))+' |')
                    if ri==0: md.append('| '+' | '.join(['---']*n)+' |')
                if wb: row+=len(data)+1
                md.append('')
            if wb:
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width=min(55,max(12,max(len(str(c.value or '')) for c in col)+2))
        if doc: doc.save(out)
        elif wb: wb.save(out)
        else: out.write_text('\n\n'.join(md).replace(' |\n\n|',' |\n|'),encoding='utf-8')
