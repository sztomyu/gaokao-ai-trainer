# -*- coding: utf-8 -*-
"""
gen_syllabus.py — 广东高考数学（新课标 I 卷）考点全解生成器
每个考点：核心内容（定义/公式/定理）+ 常考题型 + 备考提示 + 题库对应题号。
题号与《数学高考练习.docx》习题集编号一致。

用法：python gen_syllabus.py --output "广东高考数学考点全解.docx"
"""
import argparse
import datetime
import os
import re
import sys
from collections import OrderedDict

import openpyxl
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gen_exercise_book import set_cn_font, add_page_number
from gen_knowledge_map import chapter_to_module, split_kps, make_table

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE, "server", "data", "questions.xlsx")

# ================= 考点详细内容（按模块组织） =================
# match：用于从题库考点标签模糊匹配题号
MODULES = [
    {
        "name": "集合与常用逻辑用语",
        "share": "小题送分模块，通常第 1 题",
        "trend": "基础送分题，常与不等式、函数定义域结合；集合概念曾出现在 2026 压轴新定义题中，概念本质理解不可轻视。",
        "kps": [
            {"name": "集合的基本概念与运算",
             "core": "集合元素三性：确定性、互异性、无序性。交并补运算：A∩B、A∪B、∁UA。子集个数：n 元素集合有 2ⁿ 个子集、2ⁿ-1 个真子集。",
             "types": "与一元二次不等式、绝对值不等式、函数定义域/值域结合求交并补；新定义集合运算（压轴方向）。",
             "tips": "先化简再运算；注意空集讨论；数轴/韦恩图辅助。",
             "match": ["集合"]},
            {"name": "充分条件与必要条件",
             "core": "p⇒q 则 p 是 q 的充分条件，q 是 p 的必要条件；p⇔q 互为充要条件。小范围⇒大范围。",
             "types": "判断条件关系；由条件关系求参数范围（转化为集合包含）。",
             "tips": "集合观点：p 对应集合 A、q 对应 B，A⊆B ⟺ p 是 q 的充分条件。",
             "match": ["充分", "必要", "充要", "逻辑"]},
            {"name": "全称量词与存在量词",
             "core": "∀x∈M, p(x) 的否定为 ∃x∈M, ¬p(x)；∃ 的否定为 ∀。",
             "types": "命题否定书写；含量词命题真假判断；恒成立/存在性转化。",
             "tips": "否定只改量词和结论，不改条件范围。",
             "match": ["命题", "量词"]},
        ],
    },
    {
        "name": "函数与导数",
        "share": "约 18%–22%，压轴大户",
        "trend": "压轴题反套路：2026 年考构造、数形结合、逻辑推导，不考繁杂运算与偏难怪题；前两问必拿，第三问按思维过程给分。",
        "kps": [
            {"name": "函数的概念与三要素",
             "core": "定义域、值域、对应法则。求定义域：分母≠0、偶次根号内≥0、对数真数>0、tan x 中 x≠π/2+kπ。",
             "types": "求具体/抽象函数定义域；分段函数求值与求参；同一函数判定。",
             "tips": "抽象函数 f(g(x)) 定义域指 x 的范围；换元先求新元范围。",
             "match": ["定义域", "值域", "函数的概念"]},
            {"name": "函数的单调性",
             "core": "定义法（作差/作商）、图象法、复合函数同增异减、导数法 f′(x)>0 ⇒ 增。",
             "types": "判断/证明单调性；利用单调性解不等式、比大小；已知单调性求参数范围。",
             "tips": "分段函数整体单调需衔接点满足不等式；解抽象不等式先脱去 f 再注意定义域。",
             "match": ["单调"]},
            {"name": "函数的奇偶性与周期性",
             "core": "奇函数 f(-x)=-f(x) 图象关于原点对称；偶函数 f(-x)=f(x) 关于 y 轴对称。f(x+T)=f(x) 周期 T；f(x+a)=-f(x) ⇒ T=2a；关于 x=a、x=b 双对称 ⇒ T=2|a-b|。",
             "types": "奇偶性求解析式/求值；奇偶+单调解不等式；周期递推求值。",
             "tips": "奇函数在 x=0 有定义则 f(0)=0；奇偶性先看定义域是否对称。",
             "match": ["奇偶", "奇函数", "偶函数", "周期"]},
            {"name": "指数函数与对数函数",
             "core": "指数运算律 aᵐ·aⁿ=aᵐ⁺ⁿ；对数 log_a(MN)=log_a M+log_a N，换底公式 log_a b=ln b/ln a。指对数函数单调性由底数 a>1 或 0<a<1 决定。",
             "types": "指数对数运算求值；比较大小（找中间量 0/1）；指对数函数图象与性质；解指对数不等式。",
             "tips": "比大小三招：单调性、中间值、图象；注意底数分类讨论。",
             "match": ["指数函数", "对数函数", "对数运算", "指数", "对数"]},
            {"name": "函数零点与方程的根",
             "core": "零点存在定理：f 在 [a,b] 连续且 f(a)f(b)<0 ⇒ (a,b) 内有零点。零点个数 ⟺ 方程 f(x)=g(x) 根的个数 ⟺ 两图象交点个数。",
             "types": "判断零点区间；求零点个数（数形结合）；由零点个数求参数范围。",
             "tips": "分离参数 a=h(x)，研究 h 图象与水平线交点；含绝对值先分段。",
             "match": ["零点", "方程根的", "交点"]},
            {"name": "导数的几何意义与切线",
             "core": "f′(x₀) 为切线斜率；切线方程 y-f(x₀)=f′(x₀)(x-x₀)。注意『在点处切线』（该点即切点）与『过点切线』（需设切点）之别。",
             "types": "求切线方程；已知切线求参数；公切线问题；切线条数。",
             "tips": "过点 P 的切线：设切点 (t,f(t))，斜率两个表达式相等列方程。",
             "match": ["切线", "几何意义"]},
            {"name": "导数与单调性、极值、最值",
             "core": "f′(x)>0 ⇒ 增，f′(x)<0 ⇒ 减；f′(x₀)=0 且两侧变号 ⇒ 极值点（左正右负极大，反之极小）。闭区间最值 = 极值与端点值比较。",
             "types": "求单调区间（含参分类讨论）；极值/最值求参；三次函数、含 eˣ/ln x 函数。",
             "tips": "含参讨论按导函数零点个数与位置分类；单调区间端点写开区间。",
             "match": ["极值", "最值", "导数求", "导数应用", "单调区间"]},
            {"name": "恒成立与存在性问题",
             "core": "∀x, f(x)≥0 ⟺ f(x)min≥0；∃x, f(x)≥0 ⟺ f(x)max≥0。常用手段：分离参数、分类讨论、构造函数。",
             "types": "含参不等式恒成立求参数范围；双变量『任意-存在』组合。",
             "tips": "能分离参数优先分离；分离后转化为新函数最值；注意端点取等的检验。",
             "match": ["恒成立", "存在性", "参数范围"]},
            {"name": "导数证明不等式",
             "core": "构造差函数 F(x)=f(x)-g(x)，证 F(x)min>0；常用放缩：eˣ≥x+1、ln x≤x-1、ln x≤x/e 等（需会推导）。",
             "types": "证明函数不等式；数列不等式（利用函数不等式赋值）；双变量不等式（比值换元 t=x₂/x₁）。",
             "tips": "先求导判单调找最小值；放缩不等式要会独立证明，不能直接用二级结论。",
             "match": ["不等式证明", "证明"]},
        ],
    },
    {
        "name": "三角函数与解三角形",
        "share": "约 8%–12%，基础送分模块",
        "trend": "题型固定、难度较低；部分年份并入多选题压轴，注意多选『选对部分得 3 分，选错不得分』的取舍策略。",
        "kps": [
            {"name": "任意角、弧度制与诱导公式",
             "core": "180°=π rad；弧长 l=|α|r，扇形面积 S=½lr=½|α|r²。诱导公式口诀：奇变偶不变，符号看象限。同角关系：sin²α+cos²α=1，tan α=sin α/cos α。",
             "types": "弧度制换算与扇形计算；诱导公式化简求值；sin±cos 与 sin·cos 互求（平方搭桥）。",
             "tips": "(sinα+cosα)²=1+2sinαcosα，知一求二。",
             "match": ["弧度", "诱导公式", "同角"]},
            {"name": "三角恒等变换",
             "core": "和差公式 sin(α±β)、cos(α±β)、tan(α±β)；二倍角 sin2α=2sinαcosα，cos2α=cos²α-sin²α=1-2sin²α=2cos²α-1；辅助角 a sin x+b cos x=√(a²+b²)·sin(x+φ)。降幂：sin²α=(1-cos2α)/2，cos²α=(1+cos2α)/2。",
             "types": "给值求值（凑角：把所求角用已知角表示）；给值求角（缩范围定象限）；化简 f(x)=Asin(ωx+φ)。",
             "tips": "角的变换是核心：β=(α+β)-α 等；先统一角再统一名。",
             "match": ["恒等变换", "二倍角", "和差", "辅助角"]},
            {"name": "三角函数图象与性质",
             "core": "y=Asin(ωx+φ)：周期 T=2π/|ω|，振幅 |A|，对称轴 ωx+φ=π/2+kπ，对称中心 ωx+φ=kπ。图象变换：平移『左加右减针对 x』，伸缩与平移顺序影响结果。",
             "types": "由图象求解析式（A、ω、φ 三步）；单调区间/对称轴求法；ω 取值范围；图象变换。",
             "tips": "求 φ 代最值点最稳；单调区间整体代入法；平移量=|φ|/ω 易错。",
             "match": ["图象", "图象变换", "周期性", "对称"]},
            {"name": "正弦定理与余弦定理",
             "core": "正弦定理 a/sin A=b/sin B=c/sin C=2R（R 为外接圆半径）；余弦定理 a²=b²+c²-2bc cos A。面积 S=½ab sin C=½bc sin A=½ac sin B。",
             "types": "边角互化（正弦定理边化角或角化边）；求边/角/面积最值范围；判断三角形形状；实际测量应用题。",
             "tips": "边化角后用恒等变换+三角函数最值；范围问题注意 A∈(0,π) 及两边之和大于第三边。",
             "match": ["正弦定理", "余弦定理", "解三角形", "面积公式"]},
        ],
    },
    {
        "name": "平面向量",
        "share": "小题为主，约 5 分",
        "trend": "多选高频考点；常与三角、解析几何交汇；2026 卷起始题即考查向量，强调向量几何意义。",
        "kps": [
            {"name": "向量的线性运算与坐标运算",
             "core": "加减法平行四边形/三角形法则；a=(x₁,y₁), b=(x₂,y₂)：a±b=(x₁±x₂, y₁±y₂)，λa=(λx₁, λy₁)。共线定理：b∥a(a≠0) ⟺ b=λa ⟺ x₁y₂-x₂y₁=0。",
             "types": "基底表示向量（平面向量基本定理）；三点共线求参数；坐标运算。",
             "tips": "选基底后用待定系数法；三点共线转化为向量共线。",
             "match": ["线性运算", "坐标运算", "共线", "基底"]},
            {"name": "数量积与夹角",
             "core": "a·b=|a||b|cosθ=x₁x₂+y₁y₂；cosθ=(a·b)/(|a||b|)；|a|=√(x₁²+y₁²)；a⊥b ⟺ a·b=0 ⟺ x₁x₂+y₁y₂=0。投影 |a|cosθ。",
             "types": "求数量积/模/夹角；垂直平行判定；极化恒等式 a·b=¼(|a+b|²-|a-b|²) 求最值。",
             "tips": "几何意义：a·b=|a| 乘 b 在 a 方向投影；垂直条件优先坐标化。",
             "match": ["数量积", "夹角", "垂直", "向量的模"]},
            {"name": "向量与平面几何综合",
             "core": "重心 G：GA+GB+GC=0（向量形式）；外心 O：|OA|=|OB|=|OC|；垂心 H：HA·BC=HB·CA=HC·AB=0；内心 I：a·IA+b·IB+c·IC=0。",
             "types": "三角形四心的向量判定；用向量证垂直/共点；向量与解三角形综合。",
             "tips": "遇中点想中线向量 AG=(AB+AC)/3（重心）；坐标化是通法。",
             "match": ["向量与", "四心", "重心"]},
        ],
    },
    {
        "name": "数列",
        "share": "约 6%–8%，多为中档题",
        "trend": "常与不等式、函数结合；解答题步骤要规范，注意 n∈N* 的取值说明。",
        "kps": [
            {"name": "等差数列",
             "core": "通项 aₙ=a₁+(n-1)d；前 n 项和 Sₙ=n(a₁+aₙ)/2=na₁+n(n-1)d/2；等差中项 2aₙ=aₙ₋₁+aₙ₊₁；m+n=p+q ⇒ aₘ+aₙ=aₚ+a_q。",
             "types": "基本量求解（a₁、d 方程组）；Sₙ 最值（d<0 时 Sₙ 有最大值）；性质速算。",
             "tips": "Sₙ 最值：找正负分界项；{aₙ} 等差 ⟺ Sₙ=An²+Bn。",
             "match": ["等差数列"]},
            {"name": "等比数列",
             "core": "通项 aₙ=a₁qⁿ⁻¹；Sₙ=a₁(1-qⁿ)/(1-q)（q≠1），q=1 时 Sₙ=na₁；等比中项 aₙ²=aₙ₋₁aₙ₊₁；m+n=p+q ⇒ aₘaₙ=aₚa_q。",
             "types": "基本量求解；与等差混合运算；等比数列判定（aₙ₊₁/aₙ 为常数）。",
             "tips": "q 的正负与公比分类讨论勿漏 q=1；等比中项注意正负两解。",
             "match": ["等比数列"]},
            {"name": "递推数列求通项",
             "core": "累加法（aₙ₊₁=aₙ+f(n)）；累乘法（aₙ₊₁=f(n)·aₙ）；构造法 aₙ₊₁=paₙ+q ⇒ 构造等比 {aₙ+q/(p-1)}；Sₙ 与 aₙ 关系：n≥2 时 aₙ=Sₙ-Sₙ₋₁。",
             "types": "一阶线性递推；Sₙ 型递推求 aₙ；新定义递推（2026 方向）。",
             "tips": "用 Sₙ-Sₙ₋₁ 必验 n=1；构造后先求新数列通项再还原。",
             "match": ["递推", "通项"]},
            {"name": "数列求和",
             "core": "公式法；错位相减法（等差×等比型）；裂项相消法 1/[n(n+k)]=(1/k)(1/n-1/(n+k))；分组求和；倒序相加。",
             "types": "解答题第二问高频；与不等式放缩结合证 Sₙ<M。",
             "tips": "错位相减注意错位对齐与末项符号；裂项后写出前几项观察相消规律。",
             "match": ["求和", "裂项", "错位相减", "分组求和"]},
        ],
    },
    {
        "name": "立体几何",
        "share": "约 12%–16%，稳中有升",
        "trend": "题型固定套路清晰；2026 年解答题起始题位置；步骤书写完整防跳步丢分；翻折、折叠创新模型多练。",
        "kps": [
            {"name": "空间几何体的表面积与体积",
             "core": "柱体 V=Sh；锥体 V=⅓Sh；球 V=(4/3)πR³，S=4πR²；台体 V=⅓h(S₁+S₂+√(S₁S₂))。",
             "types": "直接求体积/表面积；三视图还原几何体再计算；组合体。",
             "tips": "等体积法转换底面；不规则体割补。",
             "match": ["体积", "表面积", "三视图", "几何体"]},
            {"name": "线面平行与垂直的判定",
             "core": "线面平行判定：线线平行 ⇒ 线面平行；线面垂直判定：直线垂直于平面内两相交直线。性质定理互推；面面垂直性质：一平面内垂直于交线的直线垂直于另一平面。",
             "types": "解答题第一问证明平行/垂直；位置关系判断小题（命题真假）。",
             "tips": "证明写全『因为…所以…』链条，定理条件缺一不可；找平行先找中位线/平行四边形。",
             "match": ["线面", "垂直", "平行", "位置关系"]},
            {"name": "空间角：线线角、线面角、二面角",
             "core": "向量法：线线角 cosθ=|a·b|/(|a||b|)；线面角 sinθ=|a·n|/(|a||n|)；二面角 cosθ=±(n₁·n₂)/(|n₁||n₂|)，正负由图形判断（锐/钝）。",
             "types": "解答题第二问求角；几何法（作垂线找平面角）与向量法并用。",
             "tips": "建系原则：找三条两两垂直的线；法向量设 (x,y,z) 列方程组；最后说明角的锐钝定符号。",
             "match": ["二面角", "线面角", "所成角", "空间向量"]},
            {"name": "外接球与内切球",
             "core": "长方体外接球 R=½√(a²+b²+c²)；正棱锥/柱可补形为长方体；球截面圆 r²=R²-d²（d 为球心到截面距离）。",
             "types": "求外接球半径/表面积；球与多面体组合最值。",
             "tips": "三条棱两两垂直 ⇒ 补成长方体；正四面体放入正方体。",
             "match": ["外接球", "内切球", "球"]},
            {"name": "翻折与展开问题",
             "core": "翻折前后：折痕同侧线段长度不变、角度不变；展开图用于表面最短路径（化空间为平面）。",
             "types": "平面图形翻折后判位置关系/求角求体积；表面最短路径。",
             "tips": "抓住不变量；翻折后垂直关系常用于建系。",
             "match": ["翻折", "折叠", "展开", "最短路径"]},
        ],
    },
    {
        "name": "平面解析几何",
        "share": "约 12%–15%，2026 计算量下降",
        "trend": "强调几何性质应用、多想少算；小题可用特殊值/几何性质速解；联立韦达是大题通法但要善用对称性简化。",
        "kps": [
            {"name": "直线与圆",
             "core": "直线五种方程形式；点线距 d=|Ax₀+By₀+C|/√(A²+B²)；圆标准方程 (x-a)²+(y-b)²=r²；弦长=2√(r²-d²)；切线长 √(d₀²-r²)。",
             "types": "直线与圆位置关系（d 与 r 比较）；切线/弦长/最值；圆与圆位置关系。",
             "tips": "圆的问题优先几何法（d、r），少用联立硬算。",
             "match": ["直线与圆", "圆的标准方程", "点到直线"]},
            {"name": "椭圆",
             "core": "定义 |PF₁|+|PF₂|=2a；标准方程 x²/a²+y²/b²=1(a>b>0)，c²=a²-b²，离心率 e=c/a∈(0,1)；通径 2b²/a。",
             "types": "定义求方程/最值（焦点三角形）；离心率求解（找 a,c 齐次关系）；直线与椭圆综合。",
             "tips": "焦点三角形用定义+余弦定理；e 的题先列 a、c 的齐次式。",
             "match": ["椭圆"]},
            {"name": "双曲线",
             "core": "定义 ||PF₁|-|PF₂||=2a；x²/a²-y²/b²=1，c²=a²+b²，e=c/a>1；渐近线 y=±(b/a)x；焦点到渐近线距离=b。",
             "types": "渐近线与离心率互求；定义求最值；与直线位置关系。",
             "tips": "渐近线方程令右边=0 即得；e 与渐近线斜率 k 满足 e²=1+k²。",
             "match": ["双曲线", "渐近线"]},
            {"name": "抛物线",
             "core": "定义：到焦点与准线距离相等；y²=2px 焦点 (p/2,0)，准线 x=-p/2；焦点弦 |AB|=x₁+x₂+p；通径 2p。",
             "types": "定义转化求最值（到焦点=到准线）；焦点弦性质；与直线联立。",
             "tips": "抛物线题多想定义转化，往往把折线距离拉直。",
             "match": ["抛物线"]},
            {"name": "离心率",
             "core": "椭圆 e=c/a=√(1-b²/a²)；双曲线 e=√(1+b²/a²)。求法：定义式、焦点三角形边角关系、渐近线、特殊位置。",
             "types": "小题高频：给几何条件求 e 或 e 的范围。",
             "tips": "齐次化：把条件全部换成 a、c 的齐次方程；范围题用三角形三边关系/判别式。",
             "match": ["离心率"]},
            {"name": "直线与圆锥曲线综合",
             "core": "通法：设线（斜率存在与否讨论）→ 联立消元 → 判别式 Δ>0 → 韦达定理 x₁+x₂、x₁x₂ → 代入弦长 |AB|=√(1+k²)|x₁-x₂|、中点、面积等目标式。",
             "types": "弦长与面积最值；定点定值问题；取值范围；存在性探究。",
             "tips": "2026 强调几何性质简化运算：先看对称性、几何意义再决定要不要联立；设而不求是核心思想。",
             "match": ["直线与", "弦长", "定点", "定值", "面积最值", "轨迹方程"]},
        ],
    },
    {
        "name": "概率与统计",
        "share": "约 15%–20%，分值上升",
        "trend": "情境化明显（AI 能耗、充电套餐等）；出现证明题考模型理解与逻辑表达；图表信息处理能力必考；是『送分模块』但要细心。",
        "kps": [
            {"name": "抽样方法与数字特征",
             "core": "简单随机抽样、分层抽样（按比例分配）；平均数 x̄、方差 s²=Σ(xᵢ-x̄)²/n、标准差；中位数、众数、百分位数。",
             "types": "分层抽样求各层人数；由频率分布直方图估计平均数/中位数（中位数使左右面积各半）。",
             "tips": "直方图纵轴是频率/组距，面积才是频率。",
             "match": ["抽样", "平均数", "方差", "中位数", "直方图", "数字特征"]},
            {"name": "古典概型与几何概型",
             "core": "古典概型 P(A)=A 含基本事件数/总基本事件数（等可能、有限）；几何概型 P(A)=构成 A 的度量/总度量（长度/面积/体积）。",
             "types": "列举/计数求概率；与面积长度结合的几何概型；掷骰子、抽球模型。",
             "tips": "有序无序要分清；『至少』用对立事件 P=1-P(对立)。",
             "match": ["古典概型", "几何概型"]},
            {"name": "条件概率与事件独立性",
             "core": "P(B|A)=P(AB)/P(A)；全概率公式 P(B)=ΣP(Aᵢ)P(B|Aᵢ)；贝叶斯公式；独立事件 P(AB)=P(A)P(B)。",
             "types": "条件概率计算；全概率/贝叶斯应用题（抽检、疾病筛查情境）；独立性判断。",
             "tips": "画树状图理清各分支概率；『在…条件下』即条件概率。",
             "match": ["条件概率", "独立"]},
            {"name": "离散型随机变量分布列、期望与方差",
             "core": "分布列两性质：pᵢ≥0，Σpᵢ=1；E(X)=Σxᵢpᵢ；D(X)=Σ(xᵢ-E(X))²pᵢ=E(X²)-[E(X)]²；E(aX+b)=aE(X)+b。",
             "types": "解答题大题：求分布列→期望→决策比较（方案选择）；超几何分布、二项分布模型识别。",
             "tips": "先确定 X 所有可能取值再逐个算概率；检验 Σpᵢ=1 防错。",
             "match": ["分布列", "期望", "方差"]},
            {"name": "二项分布与超几何分布",
             "core": "二项分布 X~B(n,p)：P(X=k)=C(n,k)pᵏ(1-p)ⁿ⁻ᵏ，E(X)=np，D(X)=np(1-p)；超几何分布：N 件中 M 件次品取 n 件，E(X)=nM/N。",
             "types": "有放回→二项，无放回→超几何；期望比较决策；最可能次数。",
             "tips": "模型识别是关键：独立重复试验用二项，抽样不放回用超几何。",
             "match": ["二项分布", "超几何"]},
            {"name": "正态分布",
             "core": "X~N(μ,σ²)：图象关于 x=μ 对称；P(μ-σ<X≤μ+σ)≈0.6827，2σ≈0.9545，3σ≈0.9973；对称性 P(X<μ-a)=P(X>μ+a)。",
             "types": "利用对称性求概率；3σ 原则估计/判断异常；与期望综合。",
             "tips": "全部转化为已知三段概率的对称组合。",
             "match": ["正态分布", "3σ"]},
            {"name": "回归分析与独立性检验",
             "core": "线性回归 ŷ=b̂x+â 过样本中心点 (x̄,ȳ)；相关系数 r∈[-1,1]，|r|越近 1 线性相关越强；独立性检验 K²（卡方）统计量与临界值表对照下结论。",
             "types": "给公式数据求回归方程并预测；2×2 列联表独立性检验；图表信息综合题。",
             "tips": "回归直线必过 (x̄,ȳ) 是速算点；K² 结论用『有 x% 的把握认为…有关』规范表述。",
             "match": ["回归", "独立性检验", "列联表", "相关系数"]},
        ],
    },
    {
        "name": "计数原理与二项式定理",
        "share": "小题约 5 分",
        "trend": "难度较低送分题；排列组合常与概率小题结合。",
        "kps": [
            {"name": "两个计数原理与排列组合",
             "core": "分类加法、分步乘法；排列 A(n,m)=n!/(n-m)!；组合 C(n,m)=n!/[m!(n-m)!]；C(n,m)=C(n,n-m)。常用策略：捆绑法（相邻）、插空法（不相邻）、特殊元素优先、正难则反。",
             "types": "排队/选派/分组分配问题；与概率结合的基本事件计数。",
             "tips": "先分类后分步；分组除以重复组数的阶乘；『至多至少』用排除。",
             "match": ["排列", "组合", "计数"]},
            {"name": "二项式定理",
             "core": "(a+b)ⁿ=ΣC(n,k)aⁿ⁻ᵏbᵏ；通项 T_{k+1}=C(n,k)aⁿ⁻ᵏbᵏ；系数和令 a=b=1 得 2ⁿ；奇偶项系数和各为 2ⁿ⁻¹。",
             "types": "求指定项/系数（含 xᵏ 项）；系数和（赋值法 x=1、x=-1）；常数项。",
             "tips": "写通项令指数等于目标值解 k；两个二项式相乘用『搭配』思想。",
             "match": ["二项式"]},
        ],
    },
    {
        "name": "不等式与线性规划",
        "share": "融入各题，单独小题常见",
        "trend": "一元二次不等式常与集合、函数定义域结合送分；基本不等式是小题求最值利器。",
        "kps": [
            {"name": "一元二次不等式与绝对值不等式",
             "core": "ax²+bx+c>0：a>0 时取两根之外，<0 取两根之间；|x|<a ⟺ -a<x<a，|x|>a ⟺ x>a 或 x<-a；含参讨论 Δ 与开口。",
             "types": "解不等式并集交集运算；恒成立求参（判别式法）。",
             "tips": "先化标准形（a>0）；端点能否取等单独验证。",
             "match": ["一元二次不等式", "绝对值不等式", "不等式求解"]},
            {"name": "基本不等式",
             "core": "a,b>0 时 (a+b)/2≥√(ab)，等号 ⟺ a=b；常用变形 a+b≥2√(ab)、ab≤(a+b)²/4、a²+b²≥2ab。",
             "types": "求最值（和定积最大、积定和最小）；『1 的代换』：已知 x+y=1 求 1/x+1/y 型；实际应用最优化。",
             "tips": "一正二定三相等；配凑定值；连续两次用基本不等式注意等号条件一致性。",
             "match": ["基本不等式"]},
            {"name": "线性规划",
             "core": "可行域（直线定界、特殊点定域）；目标函数 z=ax+by 的几何意义为截距族；最优解一般在顶点取得。",
             "types": "求目标函数最值/范围；含参目标函数；非线性目标（距离型、斜率型）。",
             "tips": "画图规范标出顶点坐标；斜率型 z=(y-b)/(x-a) 看作点 (x,y) 与 (a,b) 连线斜率。",
             "match": ["线性规划"]},
        ],
    },
    {
        "name": "复数",
        "share": "必考 1 小题，约 5 分",
        "trend": "题型固定计算简单，送分题必拿。",
        "kps": [
            {"name": "复数四则运算与概念",
             "core": "i²=-1；z=a+bi：实部 a 虚部 b（虚部是实数）；加减乘按多项式，除法乘共轭分母实数化 (a+bi)/(c+di)=(a+bi)(c-di)/(c²+d²)。",
             "types": "四则运算求结果；由运算结果求实部虚部/参数。",
             "tips": "除法必乘共轭；i 的幂周期 4：iⁿ 循环 i,-1,-i,1。",
             "match": ["复数的运算", "四则运算", "复数"]},
            {"name": "复数的模与共轭、复平面",
             "core": "|z|=√(a²+b²)；共轭 z̄=a-bi，z·z̄=|z|²；复平面内 z 对应点 (a,b)；|z₁-z₂| 为两点距离。",
             "types": "求模/共轭；判断象限；|z-z₀|=r 的圆轨迹。",
             "tips": "象限判断看实虚部符号；模的几何意义化归圆。",
             "match": ["复数的模", "共轭", "复平面", "象限"]},
        ],
    },
    {
        "name": "算法与其他选讲",
        "share": "偶考小题",
        "trend": "程序框图近年考查减少；坐标系与参数方程为旧选考内容，新课标卷不作大题要求，了解即可。",
        "kps": [
            {"name": "程序框图",
             "core": "三种结构：顺序、条件、循环；读图按循环逐次追踪变量值，注意先判断还是先执行（当型/直到型）。",
             "types": "读图求输出值；补全判断框条件。",
             "tips": "列表逐轮记录变量，最后一轮边界是易错点。",
             "match": ["程序框图", "框图", "算法"]},
            {"name": "坐标系与参数方程（选讲）",
             "core": "极坐标与直角坐标互化 x=ρcosθ，y=ρsinθ，ρ²=x²+y²；直线/圆/圆锥曲线参数方程；参数 t 的几何意义（直线过定点时 |t| 为距离）。",
             "types": "互化与普通方程；参数几何意义求弦长；最值范围（三角换元）。",
             "tips": "消参是主线；圆上点设 (a+rcosθ, b+rsinθ) 化三角最值。",
             "match": ["参数方程", "极坐标", "选讲"]},
        ],
    },
]


def para(doc, text="", size=10.5, bold=False, color=None, align=None,
         space_after=4, cn="宋体", indent=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    if color:
        r.font.color.rgb = RGBColor(*color)
    set_cn_font(r, cn=cn)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.3
    if indent:
        p.paragraph_format.first_line_indent = Pt(indent)
    return p, r


def load_bank_qnos():
    """加载题库数学题，返回 (按模块的 {考点标签: [题号]}, 每模块题号总数)。
    题号与习题集一致（chapter 首次出现顺序 + 读入顺序）。"""
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    ws = wb["questions"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header)}
    items = []
    for r in rows[1:]:
        if not r or not r[idx.get("qid", 0)]:
            continue
        if str(r[idx["subject"]] or "").strip() != "数学":
            continue
        items.append({
            "chapter": str(r[idx["chapter"]] or "未分章").strip() or "未分章",
            "kps": split_kps(str(r[idx["knowledge_point"]] or "")),
        })
    wb.close()
    chapters = OrderedDict()
    for it in items:
        chapters.setdefault(it["chapter"], []).append(it)
    mod_kp_qnos = {}
    mod_qcount = {}
    n = 0
    for ch, ch_items in chapters.items():
        mod = chapter_to_module(ch)
        mod_qcount[mod] = mod_qcount.get(mod, 0) + len(ch_items)
        for it in ch_items:
            n += 1
            for k in it["kps"]:
                mod_kp_qnos.setdefault((mod, k), []).append(n)
    return mod_kp_qnos, mod_qcount, n


def build(output):
    mod_kp_qnos, mod_qcount, total_q = load_bank_qnos()
    total_kps = sum(len(m["kps"]) for m in MODULES)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(fp)
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = hp.add_run("广东高考数学（新课标 Ⅰ 卷）考点全解")
    hr.font.size = Pt(9)
    hr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    set_cn_font(hr)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(10.5)
    set_cn_font(normal)

    # ---------- 封面 ----------
    for _ in range(6):
        para(doc, "")
    para(doc, "广东高考数学考点全解", size=30, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, cn="黑体", space_after=8)
    para(doc, "（新课标 Ⅰ 卷 · 逐考点详解）", size=15,
         align=WD_ALIGN_PARAGRAPH.CENTER, cn="黑体", color=(0x2F, 0x54, 0x96),
         space_after=24)
    for ln in ["覆盖：%d 大模块 / %d 个考点逐条详解" % (len(MODULES), total_kps),
               "每考点含：核心公式定理 · 常考题型 · 备考提示 · 题库对应题号",
               "题库联动：%d 道练习题（对应《数学高考练习.docx》题号）" % total_q,
               "生成日期：%s" % datetime.date.today().strftime("%Y 年 %m 月 %d 日")]:
        para(doc, ln, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    doc.add_page_break()

    # ---------- 考情总述 ----------
    h = doc.add_heading("", level=1)
    r = h.add_run("一、考情总述（新课标 Ⅰ 卷）")
    r.font.size = Pt(15); r.font.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x3B, 0x66)
    set_cn_font(r, cn="黑体")
    for txt in [
        "1. 试卷结构：广东使用全国新课标 Ⅰ 卷，满分 150 分，全卷 19 题（8 单选 + 3 多选 + 3 填空 + 5 解答），题量精简、压轴题分值提升，留足思考时间。",
        "2. 难度比例：基础 : 中档 : 难 ≈ 5 : 3 : 2，基础题占 60%–70%，确保基础题满分率是提分底线。",
        "3. 分值重心：函数与导数、解析几何、立体几何、概率统计、三角函数、数列六大模块占分超 80%。",
        "4. 命题趋势：情境化命题（科技前沿、生活场景、传统文化）、跨学科融合（约 25%–30%）、开放设问按思维过程给分、反套路反机械刷题——回归教材、理解本质是关键。",
        "5. 答题规范：解答题『见式子给分』，步骤写完整不跳步；多选题选对部分得 3 分、选错不得分，宁缺毋滥。",
    ]:
        para(doc, txt, size=10.5, space_after=6, indent=0)
    para(doc, "", space_after=2)

    # 模块一览表
    rows_ov = []
    for i, m in enumerate(MODULES, 1):
        cnt = mod_qcount.get(m["name"], 0)
        rows_ov.append(["第 %d 章　%s" % (i + 1, m["name"]), m["share"],
                        "%d 个" % len(m["kps"]), "%d 题" % cnt])
    make_table(doc, ["模块", "分值占比 / 定位", "考点数", "题库题量"], rows_ov,
               widths=[6.0, 4.6, 2.2, 2.2])
    para(doc, "注：题库题量指本题库中归入该模块的练习题数，可用于逐考点对照刷题。",
         size=9, color=(0x88, 0x88, 0x88))
    doc.add_page_break()

    # ---------- 各模块考点详解 ----------
    h = doc.add_heading("", level=1)
    r = h.add_run("二、逐模块考点详解")
    r.font.size = Pt(15); r.font.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x3B, 0x66)
    set_cn_font(r, cn="黑体")

    for i, m in enumerate(MODULES, 1):
        h2 = doc.add_heading("", level=2)
        r2 = h2.add_run("第 %d 章　%s（%s）" % (i, m["name"], m["share"]))
        r2.font.size = Pt(13.5); r2.font.bold = True
        r2.font.color.rgb = RGBColor(0x2F, 0x54, 0x96)
        set_cn_font(r2, cn="黑体")
        para(doc, "【命题趋势】%s" % m["trend"], size=10,
             color=(0x55, 0x55, 0x55), space_after=8)

        for j, kp in enumerate(m["kps"], 1):
            h3 = doc.add_heading("", level=3)
            r3 = h3.add_run("%d.%d　%s" % (i, j, kp["name"]))
            r3.font.size = Pt(11.5); r3.font.bold = True
            r3.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            set_cn_font(r3, cn="黑体")

            para(doc, "【核心内容】%s" % kp["core"], size=10.5, space_after=3)
            para(doc, "【常考题型】%s" % kp["types"], size=10.5, space_after=3)
            para(doc, "【备考提示】%s" % kp["tips"], size=10.5, space_after=3)

            # 题库题号匹配
            qnos = []
            for (mod, tag), nos in mod_kp_qnos.items():
                if mod != m["name"]:
                    continue
                if any(mk in tag for mk in kp["match"]):
                    qnos.extend(nos)
            qnos = sorted(set(qnos))
            if qnos:
                p = doc.add_paragraph()
                la = p.add_run("【题库题号】")
                la.font.size = Pt(9); la.font.bold = True
                la.font.color.rgb = RGBColor(0x1B, 0x7A, 0x43)
                set_cn_font(la)
                va = p.add_run("、".join(str(x) for x in qnos) +
                               "（共 %d 题，见《数学高考练习.docx》）" % len(qnos))
                va.font.size = Pt(9)
                va.font.color.rgb = RGBColor(0x1B, 0x7A, 0x43)
                set_cn_font(va)
                p.paragraph_format.space_after = Pt(8)
            else:
                p, _ = para(doc, "【题库题号】本题库暂无对应题目", size=9,
                            color=(0xAA, 0xAA, 0xAA), space_after=8)
        if i < len(MODULES):
            doc.add_page_break()

    doc.save(output)
    return len(MODULES), total_kps, total_q


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    m, k, q = build(args.output)
    print("OK → %s（%d 模块 / %d 考点详解 / 题库 %d 题联动）" % (args.output, m, k, q))


if __name__ == "__main__":
    main()
