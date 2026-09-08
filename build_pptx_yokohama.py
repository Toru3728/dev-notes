# -*- coding: utf-8 -*-
"""
2026年6月 介護報酬臨時改定 × 生産性向上推進体制加算【横浜市版】説明資料
となりにAI / AI×介護DX
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- palette ----------
TEAL="028090"; SEAFOAM="00A896"; MINT="02C39A"
DARK="06363E"; DARKER="042930"
INK="1F2937"; MUTED="64748B"
CARD="EAF4F5"; CARD2="F3F9F9"; WHITE="FFFFFF"; LINEGRY="D9E6E7"
HFONT="Yu Gothic UI"; BFONT="Meiryo"

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
BLANK=prs.slide_layouts[6]
def rgb(h): return RGBColor.from_string(h)

def set_font_name(run,name):
    run.font.name=name
    rPr=run._r.get_or_add_rPr()
    for tag in ("a:latin","a:ea","a:cs"):
        el=rPr.find(qn(tag))
        if el is None:
            el=rPr.makeelement(qn(tag),{}); rPr.append(el)
        el.set("typeface",name)

def add_rect(slide,x,y,w,h,fill,line=None,line_w=None,shape=MSO_SHAPE.RECTANGLE,soft=False):
    sp=slide.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h))
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb=rgb(fill)
    if line: sp.line.color.rgb=rgb(line); sp.line.width=Pt(line_w or 1)
    elif soft: sp.line.color.rgb=rgb(LINEGRY); sp.line.width=Pt(1)
    else: sp.line.fill.background()
    sp.shadow.inherit=False
    return sp

def add_oval(slide,x,y,w,h,fill,alpha=None):
    sp=slide.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(w),Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb=rgb(fill); sp.line.fill.background(); sp.shadow.inherit=False
    if alpha is not None:
        srgb=sp.fill.fore_color._xFill.find(qn('a:srgbClr'))
        a=srgb.makeelement(qn('a:alpha'),{'val':str(int(alpha*1000))}); srgb.append(a)
    return sp

def add_text(slide,x,y,w,h,runs,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,space_after=6,ls=None,wrap=True):
    tb=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=tb.text_frame
    tf.word_wrap=wrap; tf.vertical_anchor=anchor
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    for i,para in enumerate(runs):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; p.space_after=Pt(space_after); p.space_before=Pt(0)
        if ls: p.line_spacing=ls
        for (txt,size,color,bold,font,*rest) in para:
            italic=rest[0] if rest else False
            r=p.add_run(); r.text=txt; r.font.size=Pt(size); r.font.bold=bold; r.font.italic=italic
            r.font.color.rgb=rgb(color); set_font_name(r,font)
    return tb

def P(*runs): return list(runs)
def R(t,s,c,b=False,f=BFONT,i=False): return (t,s,c,b,f,i)

def title_block(s,t):
    add_text(s,0.9,0.62,11.8,0.8,[P(R(t,29,INK,True,HFONT))])
    add_rect(s,0.9,1.5,1.1,0.12,MINT)

def add_table(slide,x,y,w,data,col_w,row_h,fsize=13,header=True):
    rows=len(data); cols=len(data[0])
    tbl=slide.shapes.add_table(rows,cols,Inches(x),Inches(y),Inches(w),Inches(row_h*rows)).table
    tbl.first_row=False; tbl.horz_banding=False
    for ci,cw in enumerate(col_w): tbl.columns[ci].width=Inches(cw)
    for ri,row in enumerate(data):
        tbl.rows[ri].height=Inches(row_h)
        for ci,val in enumerate(row):
            cell=tbl.cell(ri,ci)
            cell.margin_left=Inches(0.12);cell.margin_right=Inches(0.08)
            cell.margin_top=Inches(0.02);cell.margin_bottom=Inches(0.02)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            tf=cell.text_frame; tf.word_wrap=True
            p=tf.paragraphs[0]; p.alignment=PP_ALIGN.LEFT if ci==0 else PP_ALIGN.CENTER
            r=p.add_run(); r.text=val
            ishead=header and ri==0
            r.font.size=Pt(fsize); r.font.bold=ishead or ci>=3
            r.font.color.rgb=rgb(WHITE if ishead else (TEAL if ci>=3 and not ishead else INK))
            set_font_name(r,HFONT if ishead else BFONT)
            cell.fill.solid()
            cell.fill.fore_color.rgb=rgb(TEAL if ishead else (CARD2 if ri%2==1 else WHITE))
    return tbl

# ============ 1. TITLE ============
s=prs.slides.add_slide(BLANK)
add_rect(s,0,0,13.333,7.5,DARK)
add_oval(s,10.7,-1.6,4.2,4.2,TEAL,35); add_oval(s,11.9,4.9,3.4,3.4,MINT,22); add_oval(s,10.2,3.1,1.0,1.0,SEAFOAM,60)
add_text(s,0.9,1.0,10,0.5,[P(R("となりにAI ｜ AI×介護DX 説明資料（横浜市版）",14,MINT,True,HFONT))])
add_text(s,0.9,1.95,10.5,2.6,
    [P(R("2026年6月 介護報酬 臨時改定",27,"BFE3E6",True,HFONT)),
     P(R("ICT導入が",44,WHITE,True,HFONT)),
     P(R("“報酬”に直結する時代へ",44,WHITE,True,HFONT))],space_after=4,ls=1.05)
add_text(s,0.95,5.4,11,0.6,
    [P(R("横浜市の事業所向け｜",17,"CFE9EB",False,BFONT),
       R("生産性向上推進体制加算 × 処遇改善",17,MINT,True,BFONT))])
add_text(s,0.92,6.55,11.5,0.5,[P(R("2026.06　｜　とーる（となりにAI）　｜　No Smile, No Life",12,"9DC4C8",False,BFONT))])

# ============ 2. 何が変わった（訂正版） ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"何が変わった？　— まず正しく押さえる")
# alert card
add_rect(s,0.9,1.95,11.5,1.15,CARD,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s,0.9,1.95,0.13,1.15,MINT)
add_text(s,1.3,2.1,11,0.9,
    [P(R("専用の「AI加算」は無い。",17,INK,True,HFONT),
       R("　ICTは“生産性向上推進体制加算”で評価され、その取得が処遇改善の上乗せ要件になった。",14.5,MUTED,False,BFONT))],
    anchor=MSO_ANCHOR.MIDDLE,ls=1.2)
pts=[("01","処遇改善が主役","2026年6月の臨時改定は処遇改善に特化。改定率は全体 +2.03%。"),
     ("02","ICTは加算で評価","ICT・テクノロジー導入は「生産性向上推進体制加算」で評価される。"),
     ("03","取得が“鍵”に","その加算の取得が、処遇改善加算の上位区分の要件として位置づけ。")]
cx=0.9;cw=3.78;gap=0.39;cy=3.35;ch=2.95
for i,(n,t,d) in enumerate(pts):
    x=cx+i*(cw+gap)
    add_rect(s,x,cy,cw,ch,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,x+0.35,cy+0.35,0.8,0.8,TEAL)
    add_text(s,x+0.35,cy+0.35,0.8,0.8,[P(R(n,18,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+0.35,cy+1.35,cw-0.7,0.6,[P(R(t,17,INK,True,HFONT))])
    add_text(s,x+0.35,cy+1.95,cw-0.7,0.9,[P(R(d,12.5,MUTED,False,BFONT))],ls=1.18)

# ============ 3. 加算の正体 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"加算の正体：生産性向上推進体制加算")
# two big unit cards
add_rect(s,0.9,1.95,5.6,2.05,CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,1.25,2.2,5,0.6,[P(R("加算（Ⅰ）",17,TEAL,True,HFONT))])
add_text(s,1.25,2.75,5,1.1,[P(R("100",46,INK,True,HFONT),R(" 単位/月",18,INK,False,BFONT))],anchor=MSO_ANCHOR.MIDDLE)
add_rect(s,6.8,1.95,5.6,2.05,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,7.15,2.2,5,0.6,[P(R("加算（Ⅱ）",17,SEAFOAM,True,HFONT))])
add_text(s,7.15,2.75,5,1.1,[P(R("10",46,INK,True,HFONT),R(" 単位/月",18,INK,False,BFONT))],anchor=MSO_ANCHOR.MIDDLE)
# target services
add_text(s,0.9,4.35,11.5,0.5,[P(R("対象サービス（入所・居住系が中心）",15,INK,True,HFONT))])
add_rect(s,0.9,4.85,7.6,1.55,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,1.2,5.0,7.0,1.3,
    [P(R("○ ",13,MINT,True),R("特養（介護老人福祉施設）／老健／介護医療院",13.5,INK,False,BFONT)),
     P(R("○ ",13,MINT,True),R("認知症GH／特定施設／短期入所生活介護／小多機 ほか",13.5,INK,False,BFONT))],
    anchor=MSO_ANCHOR.MIDDLE,space_after=6,ls=1.15)
add_rect(s,8.8,4.85,3.6,1.55,CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,9.1,5.0,3.0,1.3,
    [P(R("対象外",13.5,"B85042",True,HFONT)),
     P(R("訪問介護・通所介護 など",13,MUTED,False,BFONT))],anchor=MSO_ANCHOR.MIDDLE,space_after=4,ls=1.15)
add_text(s,0.9,6.55,11.5,0.4,[P(R("※ 対象範囲・要件は厚労省告示／解釈通知での確認が必要",11.5,MUTED,False,BFONT))])

# ============ 4. 横浜市だといくら ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"横浜市だと、いくら？")
add_text(s,0.9,1.75,11.5,0.55,
    [P(R("横浜市は ",15,INK,False,BFONT),R("2級地（上乗せ16%）",16,TEAL,True,HFONT),
       R("。1単位＝10円 ×（1＋0.16×人件費割合）で計算。",15,INK,False,BFONT))])
data=[["サービス例","人件費割合","1単位","加算(Ⅰ)100単位","加算(Ⅱ)10単位"],
      ["特養・短期入所生活介護","45%","10.72円","約1,072円 /月・人","約107円 /月・人"],
      ["老健・GH・特定施設・小多機","55%","10.88円","約1,088円 /月・人","約109円 /月・人"]]
add_table(s,0.9,2.45,11.5,data,[3.7,1.7,1.5,2.4,2.2],0.62,fsize=13)
add_rect(s,0.9,4.7,11.5,1.15,CARD,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s,0.9,4.7,0.13,1.15,MINT)
add_text(s,1.3,4.85,11,0.9,
    [P(R("積み上がりイメージ：",15,TEAL,True,HFONT),
       R("特養100名なら 加算(Ⅰ)で月 約10.7万円（年 約128万円）。利用者数×月数で効いてくる。",14.5,INK,False,BFONT))],
    anchor=MSO_ANCHOR.MIDDLE,ls=1.2)
add_text(s,0.9,6.05,11.5,0.85,
    [P(R("※ 金額は令和6年度地域区分（2級地16%）に基づく目安。人件費割合はサービス種別で異なる。",11.5,MUTED,False,BFONT)),
     P(R("　正確な単価・端数処理は厚労省／横浜市の最新資料で要確認。",11.5,MUTED,False,BFONT))],space_after=2,ls=1.1)

# ============ 5. 算定要件 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"算定要件：（Ⅱ）から始めて（Ⅰ）へ")
# Ⅱ column
add_rect(s,0.9,1.95,5.6,4.4,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_oval(s,1.25,2.25,0.7,0.7,SEAFOAM); add_text(s,1.25,2.25,0.7,0.7,[P(R("Ⅱ",17,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
add_text(s,2.1,2.3,4.2,0.6,[P(R("まずここから（10単位）",16,INK,True,HFONT))],anchor=MSO_ANCHOR.MIDDLE)
add_text(s,1.3,3.2,4.9,3.0,
    [P(R("・委員会の設置・開催",13.5,INK,False,BFONT)),
     P(R("・生産性向上ガイドラインに沿った改善活動",13.5,INK,False,BFONT)),
     P(R("・見守り機器等のテクノロジーを 1つ以上 導入",13.5,INK,True,BFONT)),
     P(R("・年1回、効果データをオンライン提出",13.5,INK,False,BFONT))],space_after=10,ls=1.2)
# Ⅰ column
add_rect(s,6.8,1.95,5.6,4.4,CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_oval(s,7.15,2.25,0.7,0.7,TEAL); add_text(s,7.15,2.25,0.7,0.7,[P(R("Ⅰ",17,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
add_text(s,8.0,2.3,4.2,0.6,[P(R("成果を出して上へ（100単位）",16,INK,True,HFONT))],anchor=MSO_ANCHOR.MIDDLE)
add_text(s,7.2,3.2,4.9,3.0,
    [P(R("・（Ⅱ）の要件をすべて満たす",13.5,INK,False,BFONT)),
     P(R("・データで業務改善の成果を確認",13.5,INK,False,BFONT)),
     P(R("・テクノロジーを複数（3種類）導入",13.5,INK,True,BFONT)),
     P(R("・見守り機器は全居室に設置",13.5,INK,False,BFONT)),
     P(R("・職員間の適切な役割分担",13.5,INK,False,BFONT))],space_after=8,ls=1.2)
add_text(s,0.9,6.5,11.5,0.4,[P(R("※ インカム等を使う場合、（Ⅰ）（Ⅱ）とも同じ時間帯の全介護職員が使用すること",11.5,MUTED,False,BFONT))])

# ============ 6. テクノロジーの中身 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"“テクノロジー導入”の中身（活用例）")
cards=[("見守りセンサー","夜間の転倒・離床・体調変化を検知して自動通知。夜勤の見守りを補助。"),
       ("インカム／ICT連携","職員間の連絡を効率化。同じ時間帯の全介護職員が使うのが条件。"),
       ("介護記録ソフト・AI","音声・タブレットで記録、ケアプラン作成支援。転記の手間を削減。")]
cx=0.9;cw=3.78;gap=0.39;cy=2.0;ch=3.9
for i,(t,d) in enumerate(cards):
    x=cx+i*(cw+gap)
    add_rect(s,x,cy,cw,ch,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s,x,cy,cw,0.14,MINT)
    add_oval(s,x+0.35,cy+0.5,0.85,0.85,TEAL); add_text(s,x+0.35,cy+0.5,0.85,0.85,[P(R(str(i+1),22,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+0.35,cy+1.6,cw-0.7,0.9,[P(R(t,16.5,INK,True,HFONT))],ls=1.05)
    add_text(s,x+0.35,cy+2.5,cw-0.7,1.2,[P(R(d,12.5,MUTED,False,BFONT))],ls=1.2)
add_text(s,0.9,6.3,11.5,0.5,[P(R("これらが「生産性向上推進体制加算」で評価される“テクノロジー”の具体例。種類を増やすほど（Ⅰ）に近づく。",12.5,MUTED,False,BFONT))])

# ============ 7. なぜ今 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"なぜ今、急いで取りに行くのか")
rows=[("人","深刻な人手不足","限られた人数で質を保つ手段として、ICT/テクノロジーが“必要”に。"),
      ("要","処遇改善の“上乗せ要件”","加算取得が、処遇改善加算の上位区分を取るための条件に位置づけ。"),
      ("医","医療も同時に改定","診療報酬でもAI導入が「当たり前の選択肢」として進みつつある。")]
ry=2.0;rh=1.35;rgap=0.2
for i,(ic,t,d) in enumerate(rows):
    y=ry+i*(rh+rgap)
    add_rect(s,0.9,y,11.5,rh,CARD2,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,1.25,y+0.3,0.75,0.75,SEAFOAM); add_text(s,1.25,y+0.3,0.75,0.75,[P(R(ic,18,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,2.3,y+0.18,3.6,rh-0.3,[P(R(t,18,TEAL,True,HFONT))],anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,6.0,y+0.18,6.2,rh-0.3,[P(R(d,14,INK,False,BFONT))],anchor=MSO_ANCHOR.MIDDLE,ls=1.15)

# ============ 8. 人を置いていかないDX ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,DARK)
add_oval(s,-1.4,4.6,4.0,4.0,TEAL,28); add_oval(s,11.6,-1.3,3.6,3.6,MINT,20)
add_text(s,0.9,0.75,11.5,1.0,[P(R("ただし — ",26,MINT,True,HFONT),R("「人を置いていかないDX」で。",26,WHITE,True,HFONT))])
pts=[("AIは“下書き”、最終判断は人","専門職の役割は変わらない。AIは判断を助ける道具。"),
     ("不安に寄り添う導入","「難しそう」を減らす説明と研修をセットで進める。"),
     ("個人情報はとことん慎重に","仕組み選びは安全性を最優先に。")]
py=2.15;ph=1.35;pgap=0.22
for i,(t,d) in enumerate(pts):
    y=py+i*(ph+pgap)
    add_rect(s,0.9,y,11.5,ph,DARKER,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,1.3,y+0.45,0.45,0.45,MINT)
    add_text(s,2.05,y+0.2,9.8,ph-0.35,[P(R(t,18,WHITE,True,HFONT),R("　"+d,14.5,"B7D8DB",False,BFONT))],anchor=MSO_ANCHOR.MIDDLE,ls=1.15)

# ============ 9. はじめの一歩 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"はじめの一歩（横浜市の事業所向け）")
steps=[("1","現状の棚卸し","時間のかかる業務を洗い出し、委員会を設置。"),
       ("2","テクノロジーを1つ","見守り等を1つ導入→まず加算（Ⅱ）を狙う。"),
       ("3","データ提出","年1回の効果データをオンライン提出。"),
       ("4","（Ⅰ）へ＆様式確認","種類を増やし成果を確認。横浜市の様式も確認。")]
sx0=0.9;sw=2.86;sgap=0.29;sy=2.05;sh=3.45
for i,(n,t,d) in enumerate(steps):
    x=sx0+i*(sw+sgap)
    add_rect(s,x,sy,sw,sh,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,x+sw/2-0.5,sy+0.45,1.0,1.0,TEAL); add_text(s,x+sw/2-0.5,sy+0.45,1.0,1.0,[P(R(n,30,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+0.25,sy+1.7,sw-0.5,0.7,[P(R(t,15.5,INK,True,HFONT))],align=PP_ALIGN.CENTER,ls=1.05)
    add_text(s,x+0.25,sy+2.45,sw-0.5,0.95,[P(R(d,12.5,MUTED,False,BFONT))],align=PP_ALIGN.CENTER,ls=1.18)
add_text(s,0.9,5.95,11.5,0.9,
    [P(R("※ 具体的な算定要件・単位・様式・端数処理は、本資料作成時点では要確認。",11.5,MUTED,False,BFONT)),
     P(R("　厚労省の告示・通知、横浜市の最新案内で必ず確認してください。",11.5,MUTED,False,BFONT))],space_after=2,ls=1.1)

# ============ 10. closing ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,DARK)
add_oval(s,9.8,3.6,5.2,5.2,TEAL,30); add_oval(s,11.4,-1.7,3.8,3.8,MINT,20)
add_text(s,0.95,2.4,11,2.0,[P(R("現場の笑顔を減らさないDXを、",32,WHITE,True,HFONT)),P(R("一緒に。",32,MINT,True,HFONT))],space_after=6,ls=1.1)
add_text(s,1.0,4.75,11,0.6,[P(R("となりにAI ｜ AI×介護DX 伴走パートナー　とーる",16,"CFE9EB",False,BFONT))])
add_text(s,1.0,5.45,11,0.6,[P(R("No Smile, No Life",15,MINT,True,HFONT,True))])

outdir=os.path.join("company","secretary","materials")
os.makedirs(outdir,exist_ok=True)
out=os.path.join(outdir,"介護報酬改定2026-06_生産性向上加算_横浜市版_説明資料.pptx")
prs.save(out); print("SAVED",out)
