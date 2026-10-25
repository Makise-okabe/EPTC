"""EPTC 2026 Student Innovation Challenge — 28-slide visual research presentation.

All numerical values follow the project technical report and original solver evidence.
Authoritative paper is ../../main.tex. No measured reliability/lifetime claims are made.
"""
import os
from pathlib import Path
from deck_helpers import *
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

HERE=Path(__file__).resolve().parent
OUT=HERE/"EPTC_2026_Student_Innovation_Challenge_Presentation.pptx"
prs=Presentation()
prs.slide_width=Inches(W)
prs.slide_height=Inches(H)
blank=prs.slide_layouts[6]

K=[15,21,30]
BLT=[100,250,400]
T=[[85.490,100.323,114.389],[82.625,93.344,103.544],[80.455,88.026,95.194]]
S=[[0.012094,0.012014,0.013286],[0.011960,0.011339,0.012153],[0.011866,0.010841,0.011306]]
E=[[1.2568,1.2296,1.3535],[1.2400,1.1594,1.2369],[1.2279,1.1076,1.1498]]
WAR=[[18.611,20.948,23.503],[17.778,19.012,20.496],[17.155,17.590,18.293]]
src_sim="Source: project Icepak/MAPDL numerical results; 23 Oct 2026 evidence freeze. Model-specific, not experimental data."
src_lit="Literature: Yi et al. (2025); Wen et al. (2026); Prasher (2006). Source details on references slide."

def write_note(sl,msg):
    try:
        tf=sl.notes_slide.notes_text_frame
        if tf:tf.text=msg
    except Exception:pass

def icon_circle(sl,x,y,symbol,color=TEAL):
    shp=sl.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(.44),Inches(.44))
    shp.fill.solid();shp.fill.fore_color.rgb=rgb(color);shp.line.fill.background()
    put_text(sl,x+.05,y+.09,.34,.25,symbol,size=15,color=WHITE,bold=True,align=PP_ALIGN.CENTER)

def arrow(sl,x1,y,x2,color=CORAL):
    put_line(sl,x1,y,x2,y,color,2.3)
    put_line(sl,x2-.11,y-.10,x2,y,color,2.3)
    put_line(sl,x2-.11,y+.10,x2,y,color,2.3)

def source_note(sl,s=src_sim):
    put_text(sl,.58,6.76,12.2,.20,s,size=9,color=GREY)

# 01 — COVER
s=addslide(prs,"Thermomechanical Co-Design of TIM1","FINAL COMPETITION PRESENTATION",
           notes="Introduce the challenge and the four team members. This is an engineering simulation, not experimental qualification.")
put_text(s,.75,1.40,8.20,1.20,"Interface-Aware Thermal\nInterface Material Optimization",
         size=33,color=NAVY,bold=True)
put_text(s,.78,2.85,7.95,.62,
         "A representative 700-W AI accelerator FCBGA package",
         size=20,color=GREY)
pill(s,.78,3.68,"THERMAL + MECHANICAL",TEAL,w=2.97)
pill(s,3.97,3.68,"CONTACT GEOMETRY",RUST,w=2.72)
put_text(s,.80,4.57,8.10,.43,"Fang Tianchi  ·  Ke Tangxin  ·  Du Yanzhang  ·  Duan Yihe",
         size=16,color=INK,bold=True)
put_text(s,.80,5.08,7.25,.34,"EPTC 2026 Student Innovation Challenge  |  Singapore",
         size=15,color=GREY)
card(s,9.33,1.40,3.27,4.70,fill=LIGHT)
label(s,9.60,1.78,"Key design decision",size=13)
put_text(s,9.60,2.24,2.72,.91,"100 µm\nvs 250 µm",size=29,color=NAVY,bold=True)
put_text(s,9.60,3.52,2.70,1.12,"The coolest TIM is not always the mechanically least loaded.",size=20,color=RUST,bold=True)
put_line(s,9.60,5.02,12.32,5.02,CORAL,2)
put_text(s,9.60,5.25,2.77,.50,"Matched nine-case DOE",size=14,color=INK)
source_note(s,"Source: team simulation study and the supplied EPTC 2026 competition report.")

# 02 — EXECUTIVE
s=addslide(prs,"Executive takeaway","THE ENGINEERING ANSWER",src_sim,
           "A two-candidate window emerges. No universal optimum or durability claim.")
put_text(s,.67,1.32,11.75,.61,"One package. Two valid priorities.",size=31,color=NAVY,bold=True)
stat(s,.65,2.16,3.90,2.32,"80.455 °C","Lowest modeled die temperature",TEAL,"k = 30 W/(m·K), BLT = 100 µm")
stat(s,4.70,2.16,3.90,2.32,"0.010841 MPa","Lowest sampled mean TIM stress",GREEN,"k = 30 W/(m·K), BLT = 250 µm")
stat(s,8.75,2.16,3.90,2.32,"9.46%","Stress penalty for thinner BLT",RUST,"250 → 100 µm at fixed k = 30")
card(s,.75,4.83,11.80,1.18,fill=PALE_BLUE)
put_text(s,1.03,5.04,11.22,.68,
         "Higher conductivity helps broadly; installed thickness creates the sampled thermal–mechanical trade-off.",
         size=21,color=NAVY,bold=True,align=PP_ALIGN.CENTER)

# 03 — CHALLENGE
s=addslide(prs,"Why TIM1 becomes the design bottleneck","RESEARCH PROBLEM",
           "Published TIM context: Prasher (2006); Due & Robinson (2013); Zhao et al. (2019).")
textbox_card(s,.65,1.53,3.92,3.64,"High power density",
             "A nonuniform 700-W die load creates severe thermal demands for the die-to-lid interface.",TEAL,17)
textbox_card(s,4.70,1.53,3.92,3.64,"Real interfaces are imperfect",
             "Bulk k alone misses compressed BLT, contact resistance, partial retention and local gaps.",ORANGE,17)
textbox_card(s,8.75,1.53,3.92,3.64,"Reliability is multiphysics",
             "Thinner or stiffer interfaces can change stress and strain even when warpage improves.",RUST,17)
put_text(s,.92,5.60,11.5,.55,
         "The optimization target is package behavior—not a single nominal material property.",
         size=22,bold=True,color=NAVY,align=PP_ALIGN.CENTER)

# 04 — OBJECTIVES
s=addslide(prs,"Research questions and objectives","OBJECTIVES",src_lit)
qs=[
    ("01","THERMAL","How do TIM family, k and BLT affect the modeled die temperature?"),
    ("02","INTERFACE","How much does retained contact geometry matter beyond retained area?"),
    ("03","MECHANICAL","Does the best thermal design also reduce TIM stress, strain and lid warpage?"),
    ("04","DECISION","Which sampled candidates deserve experimental qualification?")
]
for i,(n,h,desc) in enumerate(qs):
    yy=1.38+i*1.23
    card(s,.81,yy,11.70,1.05,fill=PALE if i%2 else WHITE)
    put_rect(s,.98,yy+.19,.65,.64,fill=RUST if i==3 else TEAL,radius=True)
    put_text(s,1.07,yy+.35,.47,.24,n,15,WHITE,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,1.83,yy+.14,2.02,.37,h,17,NAVY,bold=True)
    put_text(s,4.00,yy+.19,8.20,.62,desc,18,INK)

# 05 — WORKFLOW
s=addslide(prs,"An evidence-led optimization workflow","METHODS",src_sim,
           "No surrogate model is claimed. These are solved discrete scenarios with an analytical resistance interpretation.")
stages=[
 ("1","Analytical model","Bulk resistance & assumptions"),
 ("2","Icepak FEA","Materials, coverage, k–BLT DOE"),
 ("3","Numerical gates","Baseline replication & mesh checks"),
 ("4","MAPDL mechanics","One-way, fully bonded handoff"),
 ("5","Design decision","Thermal vs. stress/strain/warpage")
]
for i,(n,t,b) in enumerate(stages):
    x=.45+i*2.61
    card(s,x,1.90,2.40,2.88,fill=LIGHT)
    put_rect(s,x+.76,2.14,.88,.78,fill=TEAL if i<4 else RUST,radius=True)
    put_text(s,x+.96,2.31,.49,.36,n,size=24,color=WHITE,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,x+.18,3.10,2.08,.60,t,size=17,color=NAVY,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,x+.15,3.75,2.12,.83,b,size=14,color=INK,align=PP_ALIGN.CENTER)
    if i<4:arrow(s,x+2.42,3.26,x+2.62,ORANGE)
put_text(s,.82,5.40,11.85,.57,
         "All curves and design comparisons use solved cases—not an inferred lifetime or fabricated optimum.",
         size=18,color=RUST,bold=True,align=PP_ALIGN.CENTER)

# 06 — PACKAGE STACK
s=addslide(prs,"Representative FCBGA package model","METHODS / MODEL",
           "Geometry: study-defined die and cold plate; literature-aligned lid, TIM2 and substrate (Wen et al., 2026).")
layers=[
 ("Fixed-temperature cold plate","60 × 60 × 8 mm • top at 25 °C",RUST),
 ("TIM2","0.10 mm • k = 5.4 W/(m·K)",PALE_BLUE),
 ("Copper lid","52.4 × 52.4 × 0.70 mm","B7CCD8"),
 ("TIM1 — the design variable","k, BLT and contact area",ORANGE),
 ("Silicon die","28.5 × 28.5 × 0.70 mm • 700 W","EDC6AE"),
 ("Equivalent support / substrate","52.4 × 52.4 × 1.0 mm",PALE_BLUE)
]
for i,(t,d,c) in enumerate(layers):
    xx=1.45 if i in [0,5] else 2.1 if i in [1,2] else 3.10
    ww=10.50 if i in [0,5] else 9.20 if i in [1,2] else 7.20
    yy=1.37+i*.82
    put_rect(s,xx,yy,ww,.67,fill=c,line=LINE)
    put_text(s,xx+.15,yy+.09,ww*.48,.47,t,16,NAVY if c!=RUST else WHITE,bold=True)
    put_text(s,xx+ww*.50,yy+.11,ww*.47,.40,d,14,NAVY if c!=RUST else WHITE,align=PP_ALIGN.RIGHT)
put_text(s,.58,6.45,12.3,.25,"Schematic only • Not a proprietary accelerator geometry • One-way thermal-to-mechanical coupling",12,GREY)

# 07 — INPUT PROVENANCE
s=addslide(prs,"Model assumptions: documented, not disguised","METHODS / TRACEABILITY",
           "References: Yi et al. (2025), Wen et al. (2026), Li et al. (2023).")
for i,(title,desc,color) in enumerate([
 ("Literature-anchored","21 W/(m·K) carbon pad; 88.22% and 64.82% contact fractions",TEAL),
 ("Literature-aligned","52.4 mm substrate and lid; 0.1 mm TIM2; 60 mm cooler footprint",GREEN),
 ("Study-defined","700-W, 4×4 hotspot map; 28.5 mm die; 25 °C fixed boundary",ORANGE),
 ("Sensitivity variables","k = 15/21/30; BLT = 100/250/400 µm; carbon effective modulus/CTE",RUST)
]):
    yy=1.32+i*1.28
    card(s,.83,yy,11.65,1.10,fill=PALE if i%2 else WHITE)
    put_rect(s,1.02,yy+.18,.10,.73,fill=color)
    put_text(s,1.32,yy+.13,3.20,.42,title,17,NAVY,bold=True)
    put_text(s,4.55,yy+.17,7.54,.71,desc,17,INK)
put_text(s,1.00,6.55,11.45,.31,
         "Neither the 700 W heat map nor k = 30 W/(m·K) at 100 µm has been measured on a product.",
         size=13,color=RUST,bold=True,align=PP_ALIGN.CENTER)

# 08 — POWER MAP
s=addslide(prs,"Constructed 4 × 4 die power map","METHODS / THERMAL LOAD",src_sim)
power=[[7,42,42,7],[42,84,84,42],[42,84,84,42],[7,42,42,7]]
for r in range(4):
    for c in range(4):
        v=power[r][c];col="E6EEF2" if v==7 else "F2C39C" if v==42 else RUST
        xx=1.10+c*1.02;yy=1.48+r*1.04
        put_rect(s,xx,yy,.94,.96,fill=col)
        put_text(s,xx+.05,yy+.31,.84,.42,str(v)+" W",18,WHITE if v==84 else NAVY,
                 bold=True,align=PP_ALIGN.CENTER)
card(s,6.00,1.59,5.94,3.80,fill=LIGHT)
put_text(s,6.31,1.92,5.27,.60,"700 W total",30,TEAL,bold=True)
bullets(s,6.33,2.84,5.02,[
 "4 central tiles × 84 W = 336 W",
 "8 adjacent tiles × 42 W = 336 W",
 "4 corner tiles × 7 W = 28 W"
],size=17,gap=.68)
put_text(s,1.20,5.87,10.80,.61,
         "A fixed, nonuniform load reveals why contact loss beneath hotspots can be critical.",
         size=21,color=NAVY,bold=True,align=PP_ALIGN.CENTER)

# 09 — EQUATIONS
s=addslide(prs,"Thermal resistance: more than bulk conductivity","PHYSICAL MODEL",
           "Physical interpretation: Prasher (2006); Zhao et al. (2019). Project response metric is not JEDEC θJC.")
card(s,.69,1.45,5.75,3.17,fill=PALE)
put_text(s,.94,1.72,5.25,.46,"Intrinsic bulk contribution",18,NAVY,bold=True)
put_text(s,1.15,2.50,4.78,.72,"R_bulk ≈ BLT / (k⊥ × A)",27,TEAL,bold=True,align=PP_ALIGN.CENTER)
put_text(s,.98,3.70,5.10,.60,"Reducing BLT or increasing k lowers the ideal layer resistance.",17,INK)
card(s,6.64,1.45,5.95,3.17,fill=PALE)
put_text(s,6.89,1.72,5.38,.47,"Installed interface contribution",18,NAVY,bold=True)
put_text(s,6.87,2.48,5.49,.78,"R ≈ R_c,die + BLT/(k⊥ A_eff) + R_c,lid",
         20,RUST,bold=True,align=PP_ALIGN.CENTER)
put_text(s,6.94,3.70,5.12,.60,"Real contact adds pressure-, surface- and temperature-dependent effects.",16,INK)
card(s,1.44,5.05,10.50,1.00,fill=PALE_BLUE)
put_text(s,1.77,5.32,9.87,.38,"Reported metric:  R_eff = (peak die temperature − 25 °C) / 700 W",
         21,NAVY,bold=True,align=PP_ALIGN.CENTER)

# 10 — VALIDATION
s=addslide(prs,"Numerical credibility and verification gates","METHODS / VALIDATION",src_sim)
stat(s,.70,1.55,3.88,2.17,"0.256%","Grease medium-to-fine ΔTmax",TEAL)
stat(s,4.75,1.55,3.88,2.17,"<0.004 °C","Independent carbon reproduction",GREEN)
stat(s,8.78,1.55,3.88,2.17,"0.000031 °C","DOE control-point difference",TEAL)
card(s,.78,4.08,11.79,1.74,fill=PALE_BLUE)
put_text(s,1.08,4.33,11.23,.44,"What these checks do — and do not — prove",19,NAVY,bold=True)
put_text(s,1.10,4.92,11.14,.65,
         "They support solver consistency, not experimental validation. The later 700-W results use ideal thermal contacts.",
         size=18,color=INK,align=PP_ALIGN.CENTER)

# 11 — FAMILY SCREEN
s=addslide(prs,"Baseline comparison: grease, carbon and indium","RESULTS / MATERIAL SCREEN",
           "Project 700-W simulation. Materials have different k and BLT—this is a screen, not an isolated k effect.")
card(s,.68,1.46,8.28,4.50)
bar_h(s,.95,2.04,7.60,2.99,["Grease  •  5 / 90 µm","Carbon  •  21 / 250 µm","Indium  •  86 / 400 µm"],
      [102.534,93.344,81.883],[CORAL,TEAL,GREEN],vmax=112,
      fmt="{:.3f}",unit="°C",fs=15,label_width=2.88,value_fs=13)
card(s,9.13,1.46,3.46,4.50,fill=PALE)
put_text(s,9.43,1.81,2.93,.60,"−9.19 °C",27,TEAL,bold=True)
put_text(s,9.43,2.48,2.83,.67,"Carbon vs. grease",15,INK)
put_line(s,9.40,3.27,12.28,3.27,LINE,1)
put_text(s,9.43,3.61,2.93,.60,"−20.65 °C",27,GREEN,bold=True)
put_text(s,9.43,4.28,2.84,.60,"Indium vs. grease",15,INK)
put_text(s,9.30,5.35,3.1,.35,"Nominal ideal contact",11,RUST,bold=True)

# 12 — COVERAGE
s=addslide(prs,"Contact retention affects hotspot temperature","RESULTS / INTERFACE RETENTION",
           "Coverage fractions motivated by Yi et al. (2025). Temperature results are project idealized cases.")
card(s,.70,1.48,8.10,4.44)
bar_v(s,1.13,2.06,7.08,3.05,["100%","88.22%","64.82%"],
      [93.344,93.653,109.758],[GREEN,TEAL,RUST],vmin=75,vmax=120,unit="°",fs=15)
card(s,9.00,1.48,3.54,4.44,fill=LIGHT)
put_text(s,9.32,1.86,2.94,.39,"Important distinction",17,NAVY,bold=True)
bullets(s,9.33,2.50,2.84,[
 "Yi et al.: measured coverage ratios",
 "Here: assumed missing-contact shapes",
 "Not a measured failure progression"
],size=14,gap=.90)
put_text(s,1.10,6.23,11.18,.43,
         "The 88.22% retained-area case changes little; a more severe 64.82% ring loss changes much more.",
         18,NAVY,bold=True,align=PP_ALIGN.CENTER)

# 13 — TOPOLOGY
s=addslide(prs,"Equal coverage does not imply equal heat removal","RESULTS / CONTACT TOPOLOGY",
           "Project topology scenarios at 64.82% area; severe conductive blockage, not physical aging predictions.")
styles=["ring","edge","center"]
temps=[109.758,417.839,400.546]
for i,typ in enumerate(styles):
    x=.66+i*4.23
    card(s,x,1.43,4.03,4.55,fill=PALE)
    # Area-preserving conceptual footprints: retain exactly 64.82% by geometry.
    # These drawings are NOT images of measured TIM delamination.
    xx=x+.90; yy=1.87; side=2.22
    put_rect(s,xx,yy,side,side,fill="E7DCD8",line=LINE)
    if typ=="ring":
        retained_side=side*(0.6482**0.5)
        margin=(side-retained_side)/2
        put_rect(s,xx+margin,yy+margin,retained_side,retained_side,fill=TEAL)
    elif typ=="edge":
        put_rect(s,xx,yy,side*0.6482,side,fill=TEAL)
    else:
        put_rect(s,xx,yy,side,side,fill=TEAL)
        void_side=side*(0.3518**0.5)
        margin=(side-void_side)/2
        put_rect(s,xx+margin,yy+margin,void_side,void_side,fill="E7DCD8")
    name=["Peripheral ring","One edge band","Central void"][i]
    put_text(s,x+.22,4.31,3.61,.41,name,17,NAVY,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,x+.19,4.98,3.64,.52,"%.1f °C"%temps[i],29,GREEN if i==0 else RUST,
             bold=True,align=PP_ALIGN.CENTER)
put_text(s,.82,6.30,11.88,.32,"All three scenarios assume 64.82% retained TIM1 area; topology sketches are illustrative.",15,GREY,align=PP_ALIGN.CENTER)

# 14 — THERMAL DOE GRID
s=addslide(prs,"The full 3 × 3 conductivity–thickness thermal DOE","RESULTS / NINE SOLVED CASES",src_sim)
heatmap(s,.90,1.48,9.76,4.40,
        ["k = 15","k = 21","k = 30"],["100 µm","250 µm","400 µm"],T,fmt="{:.3f}",
        palette=("DDF1EC","D8876C"),value_size=20)
card(s,10.50,1.67,2.20,3.92,fill=LIGHT)
label(s,10.73,1.96,"Thermal best",size=10)
put_text(s,10.74,2.49,1.84,.61,"80.455 °C",24,GREEN,bold=True)
put_text(s,10.74,3.27,1.80,.55,"k = 30\n100 µm",18,NAVY,bold=True)
put_text(s,10.74,4.51,1.78,.80,"Lowest among nine sampled points only.",13,GREY)
put_text(s,1.10,6.34,11.32,.28,"Warmer colors indicate higher modeled peak die temperature; none of the grid values is experimental.",14,GREY,align=PP_ALIGN.CENTER)

# 15 — THERMAL TREND
s=addslide(prs,"The thermal trend: thin BLT and higher k both help","RESULTS / PARAMETER SENSITIVITY",src_sim)
card(s,.77,1.39,8.16,4.68)
xy_plot(s,1.04,1.91,7.57,3.63,
        [("k = 15",T[0]),("k = 21",T[1]),("k = 30",T[2])],
        [100,250,400],yrange=(75,120),xlabels=["100 µm","250 µm","400 µm"],
        yfmt="{:.0f}")
card(s,9.13,1.39,3.37,4.68,fill=LIGHT)
put_text(s,9.39,1.72,2.91,.44,"Isolated contrasts",17,NAVY,bold=True)
put_text(s,9.40,2.44,2.87,.45,"−10.719 °C",25,TEAL,bold=True)
put_text(s,9.40,2.95,2.91,.65,"250 → 100 µm\nat k = 21",14,INK)
put_line(s,9.39,3.83,12.17,3.83,LINE)
put_text(s,9.40,4.10,2.87,.45,"−5.318 °C",25,GREEN,bold=True)
put_text(s,9.40,4.63,2.91,.67,"k: 21 → 30\nat 250 µm",14,INK)
put_text(s,1.00,6.39,11.44,.29,
         "Trend is bounded to k = 15–30 W/(m·K) and BLT = 100–400 µm; not a continuous global optimum.",
         13,GREY,align=PP_ALIGN.CENTER)

# 16 — RETAINED SHORTLIST
s=addslide(prs,"Does the shortlist persist at 88.22% retention?","RESULTS / CONTACT ROBUSTNESS",src_sim)
labels=["21 / 250\nfull","21 / 250\n88.22%","30 / 250\n88.22%","30 / 100\n88.22%"]
values=[93.344,93.653,88.295,80.650]
card(s,.70,1.49,8.70,4.57)
bar_v(s,1.16,1.98,7.68,3.22,labels,values,[GREY,CORAL,TEAL,GREEN],vmin=70,vmax=100,unit="°",fs=12)
card(s,9.58,1.49,2.92,4.57,fill=PALE)
put_text(s,9.89,1.83,2.37,.55,"80.650 °C",26,GREEN,bold=True)
put_text(s,9.88,2.49,2.34,1.12,"Thin/high-k remains thermally strongest of the four tested cases.",17,INK)
put_line(s,9.88,4.05,12.17,4.05,LINE)
put_text(s,9.88,4.25,2.33,1.40,"Fixed idealized retained footprint—not a physical peel/delamination model.",13,RUST)
put_text(s,1.00,6.35,11.3,.30,"The four-case result includes recovery of an initially interrupted solver run without parameter changes.",12,GREY,align=PP_ALIGN.CENTER)

# 17 — MECH PROTOCOL
s=addslide(prs,"Thermal fields transferred into a common structural model","METHODS / MECHANICAL MODEL",src_sim)
boxes=[
 ("Icepak fields","9 independent solved\nk–BLT conditions",TEAL),
 ("Reduced handoff","Die-class and layer-\naverage temperatures",ORANGE),
 ("MAPDL quarter model","SOLID186, fully bonded\none-way response",NAVY),
 ("Compare outputs","TIM stress & strain;\nlid-top warpage",GREEN)
]
for i,(a,b,c) in enumerate(boxes):
    x=.52+i*3.22
    card(s,x,1.94,2.95,2.75,fill=LIGHT)
    put_rect(s,x+.21,2.23,2.53,.14,fill=c)
    put_text(s,x+.21,2.66,2.52,.55,a,18,NAVY,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,x+.19,3.44,2.56,.74,b,16,INK,align=PP_ALIGN.CENTER)
    if i<3:arrow(s,x+2.97,3.27,x+3.21,ORANGE)
card(s,1.12,5.29,11.12,.83,fill=PALE_BLUE)
put_text(s,1.48,5.49,10.41,.46,"Reference carbon mechanics held fixed: E = 1.0 MPa; ν = 0.49; CTE = 180 ppm/K",
         17,NAVY,bold=True,align=PP_ALIGN.CENTER)

# 18 — BRACKETS
s=addslide(prs,"Mechanical-property sensitivity: warpage is not enough","RESULTS / EFFECTIVE PROPERTIES",src_sim)
card(s,.69,1.49,7.65,4.52)
put_text(s,1.04,1.75,7.0,.37,"Warpage for two different thermal designs",17,NAVY,bold=True)
for i,(nm,a,b) in enumerate([
 ("Lower modulus",17.380,16.212),
 ("Reference modulus",19.012,17.155),
 ("Higher modulus",19.874,17.587)
]):
    yy=2.39+i*1.00
    put_text(s,.96,yy,2.00,.40,nm,14,INK)
    span=3.8
    for j,(val,c) in enumerate([(a,CORAL),(b,TEAL)]):
        ww=span*(val/22)
        y2=yy+j*.30
        put_rect(s,2.97,y2,ww,.19,fill=c)
        put_text(s,7.01,y2-.03,.88,.21,"%.2f"%val,12,NAVY,bold=True)
put_rect(s,1.11,5.48,.20,.12,fill=CORAL)
put_text(s,1.37,5.38,2.58,.31,"k = 21, 250 µm",12,INK)
put_rect(s,4.21,5.48,.20,.12,fill=TEAL)
put_text(s,4.48,5.38,2.84,.31,"k = 30, 100 µm",12,INK)
card(s,8.56,1.49,3.95,4.52,fill=LIGHT)
put_text(s,8.85,1.86,3.38,.58,"+27.25% stress",24,RUST,bold=True)
put_text(s,8.85,2.53,3.24,.49,"Higher-modulus pair",14,INK)
put_text(s,8.85,3.25,3.38,.58,"+27.32% strain",24,RUST,bold=True)
put_text(s,8.85,3.93,3.18,.70,"Warpage improved while average stress and strain worsened.",17,INK)

# 19 — MESH
s=addslide(prs,"Mesh refinement: trust the averages, not peak stress","VALIDATION / LIMITS",src_sim)
names=["Mean stress","Mean strain","Lid warpage","Local peak stress"]
changes=[.59,1.96,.038,50.2]
cols=[GREEN,TEAL,GREEN,RUST]
card(s,.75,1.46,11.85,4.74)
bar_h(s,1.09,2.08,11.00,3.45,names,changes,cols,vmax=55,
      fmt="{:.2f}",unit="%",fs=16,label_width=2.30,value_fs=15)
put_text(s,1.27,5.62,10.84,.34,"Absolute intermediate-to-fine changes; peak stress lacks mesh convergence.",
         16,RUST,bold=True,align=PP_ALIGN.CENTER)
put_text(s,1.05,6.34,11.23,.29,
         "The report does not use pointwise peak stress as a quantitative failure or fatigue criterion.",
         13,GREY,align=PP_ALIGN.CENTER)

# 20 — MATCHED DESIGN
s=addslide(prs,"A fairer mechanical comparison: nine matched cases","METHODS / DESIGN OF EXPERIMENTS",src_sim)
heatmap(s,.72,1.49,9.54,4.55,["k = 15","k = 21","k = 30"],
        ["100 µm","250 µm","400 µm"],T,fmt="{:.1f}°",value_size=19,
        palette=("E2F2ED","EFC1AF"))
card(s,10.40,1.48,2.26,4.62,fill=LIGHT)
put_text(s,10.61,1.82,1.92,.70,"9 thermal\nfields",22,NAVY,bold=True)
put_line(s,10.63,2.86,12.25,2.86,CORAL,2)
put_text(s,10.61,3.23,1.92,.72,"9 matched\nmechanical runs",21,TEAL,bold=True)
put_text(s,10.61,4.61,1.92,.86,"Same modulus, CTE, ν and structural protocol.",13,INK)
put_text(s,1.02,6.34,11.30,.31,
         "Conductivity and BLT change; each corresponding solved thermal condition feeds the same MID-property model.",
         13,GREY,align=PP_ALIGN.CENTER)

# 21 — STRESS HEATMAP
s=addslide(prs,"Average TIM stress is lowest at 250 µm","RESULTS / MATCHED MECHANICS",src_sim)
heatmap(s,.89,1.45,9.87,4.48,
        ["k = 15","k = 21","k = 30"],
        ["100 µm","250 µm","400 µm"],S,fmt="{:.6f}",value_size=18,
        palette=("D8F0E5","D68E78"))
card(s,10.43,1.61,2.17,4.25,fill=LIGHT)
label(s,10.61,1.91,"Stress minimum",size=10)
put_text(s,10.59,2.52,1.85,.63,"0.010841",21,GREEN,bold=True)
put_text(s,10.60,3.10,1.85,.27,"MPa",13,GREY)
put_text(s,10.61,3.68,1.78,.90,"k = 30\nBLT = 250 µm",15,NAVY,bold=True)
put_text(s,10.61,4.85,1.75,.62,"Within sampled, assumed model.",12,GREY)
put_text(s,.98,6.32,11.52,.30,
         "At every k in {15, 21, 30} W/(m·K), the sampled 250-µm BLT minimizes mean stress and strain.",
         15,NAVY,bold=True,align=PP_ALIGN.CENTER)

# 22 — WARPAGE & STRAIN
s=addslide(prs,"Thinner BLT improves modeled warpage","RESULTS / MATCHED MECHANICS",src_sim)
card(s,.71,1.50,8.42,4.49)
xy_plot(s,1.06,1.98,7.65,3.46,
        [("k = 15",WAR[0]),("k = 21",WAR[1]),("k = 30",WAR[2])],
        BLT,yrange=(16,24),xlabels=["100 µm","250 µm","400 µm"],
        yfmt="{:.1f}")
card(s,9.31,1.50,3.19,4.49,fill=LIGHT)
put_text(s,9.56,1.88,2.64,.49,"k = 30 W/(m·K)",18,NAVY,bold=True)
put_text(s,9.57,2.58,2.63,.55,"17.155 µm",27,TEAL,bold=True)
put_text(s,9.57,3.18,2.62,.71,"Lowest modeled lid warpage among nine cases.",15,INK)
put_line(s,9.57,4.21,12.18,4.21,LINE)
put_text(s,9.57,4.47,2.61,.86,"But mean TIM stress is not minimized at 100 µm.",16,RUST,bold=True)

# 23 — FIXED K30
s=addslide(prs,"Holding conductivity fixed reveals the trade-off","RESULTS / CONTROLLED COMPARISON",src_sim)
stat(s,.78,1.48,3.79,2.05,"−7.571 °C","Peak temperature",TEAL,"250 → 100 µm")
stat(s,4.79,1.48,3.79,2.05,"−2.47%","Lid warpage",GREEN,"250 → 100 µm")
stat(s,8.80,1.48,3.78,2.05,"+9.46%","Mean TIM stress",RUST,"250 → 100 µm")
card(s,1.17,3.91,11.01,1.78,fill=PALE_BLUE)
put_text(s,1.59,4.15,10.17,.55,"All three changes occur at fixed k = 30 W/(m·K).",22,NAVY,bold=True,align=PP_ALIGN.CENTER)
put_text(s,1.54,4.86,10.23,.46,"Mean equivalent elastic strain rises by 10.86% when BLT falls from 250 to 100 µm.",
         17,INK,align=PP_ALIGN.CENTER)

# 24 — SCATTER
s=addslide(prs,"The sampled thermal–mechanical decision map","RESULTS / MULTI-OBJECTIVE COMPARISON",src_sim)
p=[]
for i,k in enumerate(K):
    for j,b in enumerate(BLT):
        txt=("100 µm" if (k,j)==(30,0) else "250 µm" if (k,j)==(30,1) else "")
        p.append((T[i][j],S[i][j],str(k),txt))
card(s,.66,1.43,8.64,4.72)
scatter(s,.99,1.80,7.94,3.80,p,(78,118),(.0104,.0138))
card(s,9.48,1.43,3.10,4.72,fill=LIGHT)
put_text(s,9.75,1.80,2.62,.57,"Two priorities",19,NAVY,bold=True)
pill(s,9.73,2.51,"THERMAL PRIORITY",TEAL,w=2.60,small=11)
put_text(s,9.80,3.08,2.47,.49,"30 W/m·K, 100 µm",14,INK)
pill(s,9.73,3.83,"LOWER TIM STRESS",GREEN,w=2.60,small=11)
put_text(s,9.80,4.40,2.49,.49,"30 W/m·K, 250 µm",14,INK)
put_text(s,1.09,6.34,11.29,.29,"Discrete sampled operating points, not a fitted Pareto frontier or globally proven optimum.",13,GREY,align=PP_ALIGN.CENTER)

# 25 — RECOMMENDATIONS
s=addslide(prs,"Engineering recommendation: two candidates","DECISION",src_sim)
for i,(x,head,c,blt,tmp,tmpr,stress,strain,warp,trade) in enumerate([
 (.72,"THERMAL STRETCH",TEAL,"100 µm","80.455","80.650","0.011866","1.2279","17.155",
  "Choose when modeled die temperature and warpage dominate."),
 (6.77,"LOWER-STRESS",GREEN,"250 µm","88.026","88.295","0.010841","1.1076","17.590",
  "Choose when mean TIM stress and strain dominate.")
]):
    card(s,x,1.36,5.83,4.77,fill=LIGHT)
    put_rect(s,x+.13,1.52,5.57,.47,fill=c,radius=True)
    put_text(s,x+.20,1.61,5.44,.25,head,15,WHITE,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,x+.21,2.18,5.37,.48,"k = 30 W/(m·K)  |  BLT = "+blt,20,NAVY,bold=True,align=PP_ALIGN.CENTER)
    for j,(tag,val) in enumerate([
       ("Clean Tmax",tmp+" °C"),("88.22% Tmax",tmpr+" °C"),
       ("Mean stress",stress+" MPa"),("Mean strain",strain+"%"),
       ("Lid warpage",warp+" µm")
    ]):
        yy=2.87+j*.48
        put_text(s,x+.39,yy,2.44,.32,tag,14,INK)
        put_text(s,x+2.81,yy,2.49,.32,val,15,NAVY,bold=True,align=PP_ALIGN.RIGHT)
    put_text(s,x+.37,5.51,5.11,.45,trade,13,c,bold=True,align=PP_ALIGN.CENTER)
put_text(s,1.02,6.44,11.12,.20,"Different scenario models: retained-area thermal runs and fully bonded mechanical analyses are not co-simulated.",11,RUST,align=PP_ALIGN.CENTER)

# 26 — LIMITS
s=addslide(prs,"Scientific boundaries: what we cannot claim","LIMITATIONS",src_sim)
limits=[
 ("Thermal boundary","Fixed 25 °C cold-plate surface; coolant flow not resolved."),
 ("Contact-loss geometry","Illustrative ideal blockage; not measured void shapes or a damage-growth law."),
 ("Mechanical response","Assumed linear effective properties and fully bonded one-way thermal handoff."),
 ("Reliability prediction","No validated fatigue, pump-out, delamination or failure criterion; peak stress unconverged.")
]
for i,(a,b) in enumerate(limits):
    col=i%2;row=i//2;x=.79+col*6.02;y=1.55+row*2.16
    textbox_card(s,x,y,5.76,1.84,a,b,RUST if i>1 else ORANGE,size=16)
put_text(s,.95,6.38,11.47,.25,
         "Use the results for relative design screening, not measured product lifetime or manufacturing qualification.",
         13,NAVY,bold=True,align=PP_ALIGN.CENTER)

# 27 — FUTURE
s=addslide(prs,"A realistic path to reliability qualification","FUTURE WORK / INDUSTRIAL ADOPTION",
           "Qualification directions motivated by cited TIM reliability and contact-resistance literature.")
phases=[
 ("01","Material coupons","Measure through-plane k, compressed BLT, anisotropy, temperature dependence."),
 ("02","Interface tests","Measure contact resistance vs. pressure and temperature; retention maps."),
 ("03","Coupled packaging","Introduce contact pressure, deformation, interface gaps and calibrated constitutive laws."),
 ("04","Reliability tests","Thermal cycling, pump-out, adhesion, C-SAM, delamination and fatigue criteria.")
]
for i,(n,a,b) in enumerate(phases):
    y=1.41+i*1.28
    card(s,.85,y,11.65,1.12,fill=PALE if i%2 else WHITE)
    put_rect(s,1.07,y+.20,.69,.68,fill=GREEN if i==3 else TEAL,radius=True)
    put_text(s,1.14,y+.36,.54,.29,n,16,WHITE,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,1.99,y+.12,2.58,.41,a,17,NAVY,bold=True)
    put_text(s,4.75,y+.20,7.43,.71,b,16,INK)
put_text(s,1.00,6.59,11.20,.17,"No lifetime extension or industrial product readiness has been established by this simulation alone.",10,RUST,align=PP_ALIGN.CENTER)

# 28 — SUMMARY
s=addslide(prs,"Take-home messages","CONCLUSION",src_sim)
for i,(n,title,text,c) in enumerate([
 ("01","Do not optimize k in isolation","BLT, contact retention and hotspot location control the relevant thermal path.",TEAL),
 ("02","Topology matters","Equal 64.82% retained area produced very different modeled thermal outcomes.",ORANGE),
 ("03","The best thermal point is not the lowest-stress point","At k = 30, 100 µm cools best; 250 µm minimizes sampled mean stress/strain.",GREEN)
]):
    y=1.38+i*1.53
    card(s,.77,y,11.82,1.32,fill=PALE if i%2 else WHITE)
    put_rect(s,.96,y+.23,.77,.76,fill=c,radius=True)
    put_text(s,1.06,y+.43,.55,.28,n,18,WHITE,bold=True,align=PP_ALIGN.CENTER)
    put_text(s,1.94,y+.13,10.21,.45,title,20,NAVY,bold=True)
    put_text(s,1.97,y+.69,10.13,.53,text,16,INK)
put_text(s,1.20,6.19,10.92,.56,
         "Deliverable: a traceable two-candidate design window for experimental qualification.",
         21,RUST,bold=True,align=PP_ALIGN.CENTER)

# 29 — LITERATURE
s=addslide(prs,"Literature and evidence references","REFERENCES",notes="Citations are placed inline on appropriate slides. Simulated values come from the team report, not these papers.")
refs=[
"Prasher, R. S. (2006). Thermal interface materials: Historical perspective, status and future directions. Proc. IEEE. DOI: 10.1109/JPROC.2006.879796",
"Due, J., & Robinson, A. J. (2013). Reliability of thermal interface materials: A review. Applied Thermal Engineering. DOI: 10.1016/j.applthermaleng.2012.06.013",
"Zhao, J.-W. et al. (2019). Effects of surface roughness, temperature and pressure on interface thermal resistance. IJHMT. DOI: 10.1016/j.ijheatmasstransfer.2019.06.045",
"Li, J. et al. (2023). Compliance-tunable TIMs based on vertically oriented carbon fiber arrays. Composites Science and Technology. DOI: 10.1016/j.compscitech.2023.109948",
"Yi, M. et al. (2025). FCBGA package with vertical-aligned carbon-fiber thermal pad. Microelectronics Reliability. DOI: 10.1016/j.microrel.2025.115607",
"Wen, J. et al. (2026). Indium-based self-confining metal foam TIMs for large-die FCBGA. IJTS. DOI: 10.1016/j.ijthermalsci.2025.110432",
"Subramanian, V. et al. (2019). Mechanical characterization of TIMs and its challenges. Journal of Electronic Packaging. DOI: 10.1115/1.4042805",
"Du, W. et al. (2026). TIM thermal-mechanical reliability and coverage degradation prediction. Microelectronics Reliability. DOI: 10.1016/j.microrel.2026.116180",
"Yenigun, O. et al. (2026). Coupled thermal and mechanical analysis of TIMs in electronic packaging. IEEE TCPMT. DOI: 10.1109/TCPMT.2025.3633174"
]
for i,r in enumerate(refs):
    x=.74 if i<5 else 6.77
    y=1.35+(i if i<5 else i-5)*1.03
    boxw=5.89
    put_rect(s,x,y,boxw,.94,fill=PALE if i%2 else WHITE,line=LINE,radius=True)
    put_text(s,x+.16,y+.12,boxw-.28,.71,r,size=11.1,color=INK)
put_text(s,6.97,5.67,5.42,.58,
         "Project numerical source: EPTC team report and frozen Icepak/MAPDL exports; 23 Oct 2026.",
         size=12,color=NAVY,bold=True)

# 30 — BACKUP NUMERICS
s=addslide(prs,"Backup: complete nine-case comparison","APPENDIX / NUMERICAL DATA",src_sim)
xs=[.48,2.22,3.54,5.33,7.65,9.84,11.12]
widths=[1.72,1.30,1.75,2.30,2.15,1.23,1.42]
heads=["k [W/mK]","BLT [µm]","Tmax [°C]","Mean VM [MPa]","Strain [%]","Warp [µm]","Case"]
for x,w,h in zip(xs,widths,heads):
    put_rect(s,x,1.42,w,.55,fill=NAVY)
    put_text(s,x+.02,1.60,w-.06,.30,h,11.3,WHITE,bold=True,align=PP_ALIGN.CENTER)
idx=0
for i,k in enumerate(K):
    for j,b in enumerate(BLT):
        y=2.05+idx*.45
        if idx%2==0:put_rect(s,.49,y-.025,12.03,.43,fill=PALE)
        vals=[str(k),str(b),"%.3f"%T[i][j],"%.6f"%S[i][j],"%.4f"%E[i][j],"%.3f"%WAR[i][j],
              "LOWEST T" if (i,j)==(2,0) else "LOWEST σ" if (i,j)==(2,1) else ""]
        for x,w,val in zip(xs,widths,vals):
            put_text(s,x+.04,y+.03,w-.10,.31,val,12.2,
                     GREEN if (i,j) in [(2,0),(2,1)] else INK,
                     bold=(i,j) in [(2,0),(2,1)],align=PP_ALIGN.CENTER)
        idx+=1

# Slide-level geometry QA.
issues=[]
for si,sl in enumerate(prs.slides,1):
    for sh in sl.shapes:
        x=sh.left/Inches(1);y=sh.top/Inches(1)
        w=sh.width/Inches(1);h=sh.height/Inches(1)
        if x<-.03 or y<-.03 or x+w>W+.11 or y+h>H+.11:
            issues.append((si,round(x,3),round(y,3),round(w,3),round(h,3)))
if issues:
    raise RuntimeError("Out-of-bounds elements: "+repr(issues[:30]))
assert len(prs.slides)==30, len(prs.slides)
prs.save(str(OUT))
print("CREATED",OUT)
print("SLIDE_COUNT",len(prs.slides))
print("FILE_BYTES",OUT.stat().st_size)
print("Geometry checks PASS")
