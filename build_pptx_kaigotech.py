# -*- coding: utf-8 -*-
"""
介護テクノロジー導入支援事業（令和8年度）活用提案資料
となりにAI / AI×介護DX
ソース: company/secretary/ideas/介護テクノロジー導入支援事業_まとめ.md（2026-07-07時点）
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
AMBER="D97706"; AMBERBG="FDF3E3"
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

def footnote(s,lines,y=6.55):
    add_text(s,0.9,y,11.5,0.85,[P(R(t,11.5,MUTED,False,BFONT)) for t in lines],space_after=2,ls=1.1)

# ============ 1. TITLE ============
s=prs.slides.add_slide(BLANK)
add_rect(s,0,0,13.333,7.5,DARK)
add_oval(s,10.7,-1.6,4.2,4.2,TEAL,35); add_oval(s,11.9,4.9,3.4,3.4,MINT,22); add_oval(s,10.2,3.1,1.0,1.0,SEAFOAM,60)
add_text(s,0.9,1.0,10,0.5,[P(R("となりにAI ｜ AI×介護DX ご提案資料",14,MINT,True,HFONT))])
add_text(s,0.9,1.95,11.2,2.7,
    [P(R("介護テクノロジー導入支援事業（令和8年度）",25,"BFE3E6",True,HFONT)),
     P(R("補助金を使って、",44,WHITE,True,HFONT)),
     P(R("AI・ICT導入の負担を最小に。",44,WHITE,True,HFONT))],space_after=4,ls=1.05)
add_text(s,0.95,5.4,11.5,0.6,
    [P(R("最大 補助率3/4・パッケージ型なら4/5 ｜ ",17,"CFE9EB",False,BFONT),
       R("AIケアプラン支援ソフトも対象",17,MINT,True,BFONT))])
add_text(s,0.92,6.55,11.5,0.5,[P(R("2026.07　｜　とーる（となりにAI）　｜　No Smile, No Life",12,"9DC4C8",False,BFONT))])

# ============ 2. 制度の仕組み ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"制度の仕組み　— 国の基金 × 都道府県の補助金")
add_rect(s,0.9,1.95,11.5,1.15,CARD,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s,0.9,1.95,0.13,1.15,MINT)
add_text(s,1.3,2.1,11,0.9,
    [P(R("国が2/3を負担する基金事業。",17,INK,True,HFONT),
       R("　実施主体は都道府県。だから補助率・時期・上限額は地域ごとに違う。",14.5,MUTED,False,BFONT))],
    anchor=MSO_ANCHOR.MIDDLE,ls=1.2)
pts=[("国","財源は国の基金","国が2/3を負担する基金を通じて、介護テクノロジー導入を後押しする国策。"),
     ("県","実施は都道府県","都道府県が実施主体。申請窓口・要件・時期は県ごとに設定される。"),
     ("！","地域で条件が違う","補助率・上限・募集時期は県ごとに異なる。提案先の地域を都度確認。")]
cx=0.9;cw=3.78;gap=0.39;cy=3.35;ch=2.95
for i,(n,t,d) in enumerate(pts):
    x=cx+i*(cw+gap)
    add_rect(s,x,cy,cw,ch,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,x+0.35,cy+0.35,0.8,0.8,TEAL if i<2 else AMBER)
    add_text(s,x+0.35,cy+0.35,0.8,0.8,[P(R(n,18,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+0.35,cy+1.35,cw-0.7,0.6,[P(R(t,17,INK,True,HFONT))])
    add_text(s,x+0.35,cy+1.95,cw-0.7,0.95,[P(R(d,12.5,MUTED,False,BFONT))],ls=1.18)
footnote(s,["※ 本資料の数値は2026年7月時点の公開情報に基づく目安。年度・都道府県で変わるため、申請前に最新要領を必ず確認。"])

# ============ 3. 補助率 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"補助率　— 費用の大半を補助でカバーできる")
add_rect(s,0.9,1.95,5.6,2.75,CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,1.25,2.2,5,0.6,[P(R("通常導入",17,TEAL,True,HFONT))])
add_text(s,1.25,2.8,5,1.2,[P(R("最大 3/4",46,INK,True,HFONT))],anchor=MSO_ANCHOR.MIDDLE)
add_text(s,1.25,4.05,4.9,0.6,[P(R("都道府県が 3/4 または 1/2 を上限に設定（2段階方式）",12.5,MUTED,False,BFONT))],ls=1.15)
add_rect(s,6.8,1.95,5.6,2.75,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,7.15,2.2,5,0.6,[P(R("パッケージ型導入（複数機器まとめて）",16,SEAFOAM,True,HFONT))])
add_text(s,7.15,2.8,5,1.2,[P(R("4/5",46,INK,True,HFONT),R("（80%）",20,INK,False,BFONT))],anchor=MSO_ANCHOR.MIDDLE)
add_text(s,7.15,4.05,4.9,0.6,[P(R("基準額 1,000万円。大規模なまとめ導入に有利",12.5,MUTED,False,BFONT))],ls=1.15)
add_rect(s,0.9,5.05,11.5,1.15,AMBERBG,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s,0.9,5.05,0.13,1.15,AMBER)
add_text(s,1.3,5.2,11,0.9,
    [P(R("要チェック：",15,AMBER,True,HFONT),
       R("令和7年度は「第三者による業務改善支援」の実施が補助要件だった。令和8年度も同様の可能性あり（要確認）。",14,INK,False,BFONT))],
    anchor=MSO_ANCHOR.MIDDLE,ls=1.2)
footnote(s,["※ 補助率・基準額は都道府県の交付要綱により異なる。提案先の所在都道府県の最新要綱を確認のこと。"],y=6.6)

# ============ 4. 対象機器・対象経費 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"何が対象？　— AIケアプラン支援ソフトも対象経費")
add_text(s,0.9,1.8,11.5,0.5,[P(R("重点支援対象（業務時間削減の効果が確認済み）",15,INK,True,HFONT))])
items=[("記録","介護記録ソフト","記録業務のデジタル化。音声入力・AI整形と組み合わせて効果大"),
       ("見守","見守り機器","夜間巡視の負担軽減。センサーで状態を把握"),
       ("通話","インカム","職員間の連携をリアルタイム化。移動と伝達のムダを削減")]
cx=0.9;cw=3.78;gap=0.39;cy=2.35;ch=2.0
for i,(n,t,d) in enumerate(items):
    x=cx+i*(cw+gap)
    add_rect(s,x,cy,cw,ch,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,x+0.3,cy+0.3,0.65,0.65,TEAL)
    add_text(s,x+0.3,cy+0.3,0.65,0.65,[P(R(n,11,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+1.1,cy+0.32,cw-1.35,0.6,[P(R(t,15.5,INK,True,HFONT))],anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+0.3,cy+1.1,cw-0.6,0.85,[P(R(d,11.5,MUTED,False,BFONT))],ls=1.15)
add_rect(s,0.9,4.6,11.5,1.15,CARD,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s,0.9,4.6,0.13,1.15,MINT)
add_text(s,1.3,4.75,11,0.9,
    [P(R("AIケアプラン原案作成支援ソフトも対象経費。",16,TEAL,True,HFONT),
       R("　ほかに移乗支援・入浴支援機器など。対象はTAIS（福祉用具情報システム）掲載品がベース。",13.5,MUTED,False,BFONT))],
    anchor=MSO_ANCHOR.MIDDLE,ls=1.2)
footnote(s,["※ 対象機器・経費の範囲は都道府県要綱およびTAISカタログで最終確認。"],y=6.1)

# ============ 5. 補助上限額の例 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"いくら出る？　— 補助上限額の例")
add_rect(s,0.9,1.95,7.3,4.25,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,1.25,2.2,6.6,0.5,[P(R("介護記録ソフトの場合",16,TEAL,True,HFONT))])
add_text(s,1.25,2.75,6.6,1.15,[P(R("100〜250",44,INK,True,HFONT),R(" 万円",18,INK,False,BFONT))],anchor=MSO_ANCHOR.MIDDLE)
add_text(s,1.25,3.95,6.7,2.1,
    [P(R("・職員数に応じて変動（変動なしの場合は一律250万円）",13,INK,False,BFONT)),
     P(R("・端末・Wi-Fi整備込みなら ＋15万円",13,INK,False,BFONT)),
     P(R("・ケアプランデータ連携（5事業所以上）で ＋5万円",13,INK,False,BFONT))],space_after=9,ls=1.2)
add_rect(s,8.5,1.95,3.9,4.25,CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,8.85,2.2,3.2,0.5,[P(R("パッケージ型",16,SEAFOAM,True,HFONT))])
add_text(s,8.85,2.75,3.2,1.15,[P(R("最大1,000",34,INK,True,HFONT),R(" 万円",16,INK,False,BFONT))],anchor=MSO_ANCHOR.MIDDLE)
add_text(s,8.85,3.95,3.3,2.1,
    [P(R("複数機器のまとめ導入",13,INK,False,BFONT)),
     P(R("＋ 業務改善支援",13,INK,False,BFONT)),
     P(R("45〜48万円 別枠",14,TEAL,True,HFONT))],space_after=8,ls=1.2)
footnote(s,["※ 金額は2026年7月時点の国スキームの例。実際の上限は都道府県の交付要綱で確認。"])

# ============ 6. 段階導入の提案 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"ご提案　— 小さく始めて、補助金で広げる")
add_text(s,0.9,1.8,11.5,0.5,
    [P(R("いきなり全部はやらない。",15,INK,True,HFONT),
       R("現場が慣れる順番で、補助金を最大限使いながら段階導入する。",14,MUTED,False,BFONT))])
steps=[("1","まず「記録」から","音声入力＋AI整形で記録業務を軽く。重点支援対象で補助も通りやすい。現場の成功体験づくり。"),
       ("2","見守り・インカム","夜間・連携の負担を削減。効果データが取りやすく、加算・次年度申請の実績にもなる。"),
       ("3","AIケアプラン支援","AIケアプラン原案作成支援ソフトを導入。まとめ導入ならパッケージ型4/5も視野に。")]
sx0=0.9;sw=3.6;sgap=0.55;sy=2.5;sh=3.55
for i,(n,t,d) in enumerate(steps):
    x=sx0+i*(sw+sgap)
    add_rect(s,x,sy,sw,sh,CARD2 if i<2 else CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,x+sw/2-0.5,sy+0.4,1.0,1.0,TEAL)
    add_text(s,x+sw/2-0.5,sy+0.4,1.0,1.0,[P(R(n,30,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+0.25,sy+1.6,sw-0.5,0.6,[P(R(t,16.5,INK,True,HFONT))],align=PP_ALIGN.CENTER)
    add_text(s,x+0.3,sy+2.25,sw-0.6,1.2,[P(R(d,12,MUTED,False,BFONT))],ls=1.2)
    if i<2:
        add_text(s,x+sw+0.05,sy+1.3,0.5,0.8,[P(R("→",22,MINT,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
footnote(s,["※ ステップ1〜2の効果データは「生産性向上推進体制加算」の要件（効果測定・データ提出）にもそのまま活きる。"],y=6.35)

# ============ 7. 神奈川県の状況 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"神奈川県の状況（2026年7月時点）")
rows=[("済","伴走支援の募集は終了","令和8年度「伴走支援」対象施設の募集は 5/1〜5/20 で実施済み。",TEAL),
      ("？","当初補助金は時期未確定","令和8年度当初の通常補助金の申請期間は現時点で未公表 → 公表され次第すぐ動けるよう準備。",AMBER)]
cy=1.95
for i,(n,t,d,c) in enumerate(rows):
    y=cy+i*1.5
    add_rect(s,0.9,y,11.5,1.3,CARD2 if i==0 else AMBERBG,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,1.25,y+0.3,0.7,0.7,c)
    add_text(s,1.25,y+0.3,0.7,0.7,[P(R(n,16,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,2.2,y+0.18,9.9,0.5,[P(R(t,16,INK,True,HFONT))])
    add_text(s,2.2,y+0.68,9.9,0.55,[P(R(d,12.5,MUTED,False,BFONT))],ls=1.15)
add_text(s,0.9,5.15,11.5,0.5,[P(R("相談窓口",15,INK,True,HFONT))])
add_rect(s,0.9,5.65,5.6,1.05,CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,1.2,5.78,5.0,0.8,
    [P(R("横浜市総合リハビリテーションセンター",13,INK,True,BFONT)),
     P(R("介護ロボット相談窓口",12,MUTED,False,BFONT))],space_after=3,ls=1.1)
add_rect(s,6.8,5.65,5.6,1.05,CARD,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s,7.1,5.78,5.1,0.8,
    [P(R("かながわ福祉サービス振興会",13,INK,True,BFONT)),
     P(R("carerobot.kanafuku.jp",12,TEAL,False,BFONT))],space_after=3,ls=1.1)

# ============ 8. となりにAIの伴走支援 ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,WHITE)
title_block(s,"となりにAIがやること　— 申請から定着まで伴走")
steps=[("聞く","現状ヒアリング","現場の困りごとを整理。どの業務が一番重いかを一緒に見える化"),
       ("選ぶ","機器・補助金の選定","TAIS掲載品から現場に合う機器を選び、使える補助メニューを特定"),
       ("出す","申請サポート","要綱の読み込み・必要書類の準備・効果測定計画づくりを支援"),
       ("育つ","導入・定着伴走","職員研修と使い方フォロー。「置いていかれる人」を作らない")]
sx0=0.9;sw=2.72;sgap=0.2;sy=2.1;sh=3.7
for i,(n,t,d) in enumerate(steps):
    x=sx0+i*(sw+sgap)
    add_rect(s,x,sy,sw,sh,CARD2,soft=True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_oval(s,x+sw/2-0.5,sy+0.4,1.0,1.0,TEAL)
    add_text(s,x+sw/2-0.5,sy+0.4,1.0,1.0,[P(R(n,17,WHITE,True,HFONT))],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    add_text(s,x+0.2,sy+1.6,sw-0.4,0.6,[P(R(t,14.5,INK,True,HFONT))],align=PP_ALIGN.CENTER,ls=1.05)
    add_text(s,x+0.25,sy+2.3,sw-0.5,1.25,[P(R(d,11.5,MUTED,False,BFONT))],ls=1.2)
footnote(s,["※ 令和7年度要件だった「第三者による業務改善支援」にも、外部支援者として対応可能（8年度要件は要確認）。"],y=6.2)

# ============ 9. closing ============
s=prs.slides.add_slide(BLANK); add_rect(s,0,0,13.333,7.5,DARK)
add_oval(s,9.8,3.6,5.2,5.2,TEAL,30); add_oval(s,11.4,-1.7,3.8,3.8,MINT,20)
add_text(s,0.95,2.2,11.5,2.0,
    [P(R("補助金は「安く買う」ためじゃなく、",30,WHITE,True,HFONT)),
     P(R("現場が笑顔で使いこなすための時間を買うもの。",30,MINT,True,HFONT))],space_after=6,ls=1.15)
add_text(s,1.0,4.75,11,0.6,[P(R("となりにAI ｜ AI×介護DX 伴走パートナー　とーる",16,"CFE9EB",False,BFONT))])
add_text(s,1.0,5.45,11,0.6,[P(R("No Smile, No Life",15,MINT,True,HFONT,True))])
add_text(s,1.0,6.5,11.5,0.5,
    [P(R("本資料の数値は2026年7月時点。都道府県・年度により変わります。申請前に最新の交付要綱をご確認ください。",11,"9DC4C8",False,BFONT))])

outdir=os.path.join("company","secretary","materials")
os.makedirs(outdir,exist_ok=True)
out=os.path.join(outdir,"介護テクノロジー導入支援事業_提案資料.pptx")
prs.save(out); print("SAVED",out)
