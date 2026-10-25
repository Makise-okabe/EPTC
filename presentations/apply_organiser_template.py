"""Apply the ORIGINAL organiser PPTX artwork to the editable EPTC research slides.

Required inputs:
  presentations/organiser_template.pptx  -- untouched organiser-provided slide
  presentations/EPTC_2026_Student_Innovation_Challenge_Presentation.pptx
Output:
  presentations/EPTC_2026_Official_Template_Presentation.pptx

This is deliberately a *second* output. The previous presentation is preserved.
The organisers' slide master, layout, native graphics and footer are inherited
or copied from the original presentation; they are not manually redrawn.
"""
from __future__ import annotations

from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "organiser_template.pptx"
OLD = HERE / "EPTC_2026_Student_Innovation_Challenge_Presentation.pptx"
OUTPUT = HERE / "EPTC_2026_Official_Template_Presentation.pptx"

TOP_MARGIN = .78
BOTTOM_START = 7.10

def copy_shape_with_relations(sh, src_slide, dst_slide):
    """Clone the original DrawingML, remapping image and hyperlink rIds."""
    el = deepcopy(sh.element)
    rattrs = (qn("r:embed"), qn("r:link"), qn("r:id"))
    mapped = {}
    for child in el.iter():
        for a in rattrs:
            rid = child.get(a)
            if not rid:
                continue
            if rid not in src_slide.part.rels:
                continue
            if rid not in mapped:
                rel = src_slide.part.rels[rid]
                mapped[rid] = dst_slide.part.relate_to(
                    rel.target_ref if rel.is_external else rel.target_part,
                    rel.reltype,
                    is_external=rel.is_external,
                )
            child.set(a, mapped[rid])
    dst_slide.shapes._spTree.insert_element_before(el, 'p:extLst')
    return dst_slide.shapes[-1]

def title_of_slide(slide):
    possible = []
    for sh in slide.shapes:
        if not sh.has_text_frame or not sh.text.strip():
            continue
        text = sh.text.strip()
        if sh.top < Inches(.84) and sh.width >= Inches(4):
            if text.lower() not in {"eptc","2026","title"}:
                possible.append((len(text), text))
    return max(possible)[1] if possible else "TIM1 thermal–mechanical co-design"

def set_title_preserve_format(slide, title):
    """Replace existing Title in source template rather than cover artwork."""
    possible = [sh for sh in slide.shapes
                if sh.has_text_frame and
                sh.text.strip().lower() == "title"]
    if not possible:
        # Some official templates place the title in a layout placeholder.
        possible = [sh for sh in slide.shapes
                    if sh.has_text_frame and sh.top < Inches(.9)
                    and sh.width >= Inches(4)
                    and "eptc" not in sh.text.lower()]
    if not possible:
        raise RuntimeError("Cannot find the original template title shape; refuse to draw over official artwork")
    sh = max(possible, key=lambda z: z.width)
    paragraph = sh.text_frame.paragraphs[0]
    if not paragraph.runs:
        paragraph.text = title
        paragraph.font.bold = True
        paragraph.font.color.rgb = RGBColor(255,255,255)
    else:
        paragraph.runs[0].text = title
        for run in paragraph.runs[1:]:
            run.text = ""
        main = paragraph.runs[0].font
        main.name = "Aptos Display"
        main.size = Pt(20 if len(title)>68 else 23 if len(title)>52 else 27)
        main.bold = True
        main.color.rgb = RGBColor(255,255,255)
    paragraph.alignment = PP_ALIGN.CENTER
    sh.text_frame.word_wrap = False

def update_footer_number(slide,index):
    """Keep native footer text style; update the slide-specific page number."""
    candidates = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text.strip() == "1" and sh.top > Inches(6.95):
            candidates.append(sh)
    if candidates:
        for sh in candidates:
            for para in sh.text_frame.paragraphs:
                if para.runs:
                    para.runs[0].text = str(index)
                    for run in para.runs[1:]:run.text = ""
                else: para.text = str(index)
    else:
        # A template may have footer number supplied by slide layout/master.
        # In that case add only the numeral, not a recreated footer.
        from pptx.enum.text import MSO_ANCHOR
        tb=slide.shapes.add_textbox(Inches(12.69),Inches(7.14),Inches(.32),Inches(.26))
        tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
        tb.text_frame.margin_left=Inches(0)
        tb.text_frame.margin_right=Inches(0)
        p=tb.text_frame.paragraphs[0]
        p.text=str(index)
        p.alignment=PP_ALIGN.RIGHT
        for run in p.runs:
            run.font.name="Arial"
            run.font.size=Pt(11)
            run.font.color.rgb=RGBColor(255,255,255)

def collect_template_images(slide):
    return {sha256(sh.image.blob).hexdigest()
            for sh in slide.shapes if sh.shape_type == 13}

def main():
    if not TEMPLATE.exists():
        raise FileNotFoundError(
            "Upload the organiser's untouched original PPTX to "
            "presentations/organiser_template.pptx (GitHub Add file > Upload files)."
        )
    if not OLD.exists():
        raise FileNotFoundError(f"Research deck missing: {OLD.name}")
    final=Presentation(TEMPLATE)
    earlier=Presentation(OLD)
    assert len(final.slides)==1, "Expect exactly the organiser's one-slide template"
    assert len(earlier.slides)>=30, "Existing competition research deck seems incomplete"

    exemplar=final.slides[0]
    original_sha={sha256(sh.image.blob).hexdigest()
                  for sh in exemplar.shapes if sh.shape_type==13}
    if not (exemplar.shapes or exemplar.slide_layout.shapes):
        raise RuntimeError("Official PPTX has no visible shapes or slide layout")

    prior_sh=list(exemplar.shapes)
    copied=0
    for idx,src in enumerate(earlier.slides,1):
        if idx==1:
            out=exemplar
        else:
            out=final.slides.add_slide(exemplar.slide_layout)
            for sh in prior_sh:
                copy_shape_with_relations(sh,exemplar,out)
        set_title_preserve_format(out,title_of_slide(src))
        # Copy only RESEARCH CONTENT, never the homemade header/footer.
        for sh in src.shapes:
            top=sh.top / Inches(1)
            center_y=(sh.top+sh.height/2)/Inches(1)
            if top < TOP_MARGIN-.005 or center_y>=BOTTOM_START:
                continue
            new=copy_shape_with_relations(sh,src,out)
            # Preserve content relative to the original slide surface if
            # the supplied layout has a different physical size.
            if abs(final.slide_width-earlier.slide_width)>Inches(.02):
                scale_x=final.slide_width/earlier.slide_width
                new.left=int(new.left*scale_x)
                new.width=int(new.width*scale_x)
            if abs(final.slide_height-earlier.slide_height)>Inches(.02):
                scale_y=final.slide_height/earlier.slide_height
                new.top=int(new.top*scale_y)
                new.height=int(new.height*scale_y)
            copied+=1
        update_footer_number(out,idx)

    assert len(final.slides)==len(earlier.slides)
    final.save(OUTPUT)
    check=Presentation(OUTPUT)
    assert len(check.slides)==len(earlier.slides)
    expect=["80.455","0.010841","9.46%","88.22%","64.82%","Yi, M.","Wen, J."]
    text="\n".join(sh.text for sl in check.slides for sh in sl.shapes if sh.has_text_frame)
    for value in expect:
        if value not in text:raise ValueError("Missing required content: "+value)
    # Verify template's original image data survived losslessly in first slide.
    output_image_sha=collect_template_images(check.slides[0])
    assert original_sha.issubset(output_image_sha), "Original organiser media changed"
    for idx,sl in enumerate(check.slides,1):
        first_text=[sh.text for sh in sl.shapes if sh.has_text_frame]
        assert not any(t.strip()=="Title" for t in first_text), f"Title placeholder remains on {idx}"
    print("SUCCESS: official template transferred without redrawing branding")
    print("SLIDES:",len(check.slides),"CONTENT SHAPES:",copied,
          "TEMPLATE_IMAGE_HASHES:",len(original_sha))
    print("OUTPUT:",OUTPUT, "BYTES:",OUTPUT.stat().st_size)

if __name__ == "__main__":
    main()
