# -*- coding: utf-8 -*-
import streamlit as st
import json, io
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from docx import Document

# 1. 页面配置
st.set_page_config(page_title="全能 AI 化学教研工作站", page_icon="🧪", layout="wide")

st.title("🧪 全能 AI 高中化学教研工作站")
st.caption("一键自动生成：16:9 高颜值多栏PPT课件 + Word标准化教案 + 高考级随堂练习题")

# 侧边栏配置
st.sidebar.header("⚙️ API 与参数设置")
api_key = st.sidebar.text_input("DeepSeek API Key", value="sk-45096723d48b4687a215456ae97f30b3", type="password")
base_url = st.sidebar.text_input("Base URL", value="https://api.deepseek.com")

# 主界面输入
col_input1, col_input2 = st.columns([2, 1])
with col_input1:
    topic = st.text_input("📌 化学课题名称：", "氧化还原反应")
with col_input2:
    grade = st.selectbox("🎯 适用年级/模块：", [
        "高中化学必修一", 
        "高中化学必修二", 
        "选择性必修一（化学反应原理）", 
        "选择性必修三（有机化学）"
    ])

# 自定义补充要求输入框
custom_requirements = st.text_area(
    "💡 补充个人教学要求 / 必讲内容（可选）：", 
    placeholder="在此输入你想加入的个人内容，例如：\n1. 重点强调双线桥法和单线桥法的区别\n2. 包含守恒规律的计算公式与解题口诀\n3. 结合高考真题常见易错陷阱...",
    height=100
)

# --- 高级 PPT 绘制引擎 ---
def build_ppt(topic, grade, slides_data):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9
    
    # 色彩体系
    C_PRIMARY = RGBColor(15, 23, 42)    # 深色主色
    C_ACCENT = RGBColor(234, 88, 12)    # 活力橙
    C_BG = RGBColor(248, 250, 252)        # 浅灰背景
    C_CARD_BG = RGBColor(255, 255, 255)   # 纯白卡片
    C_TEXT_MAIN = RGBColor(30, 41, 59)  # 主文本色
    
    # 1. 封面页
    slide_cover = prs.slides.add_slide(prs.slide_layouts[6])
    slide_cover.background.fill.solid()
    slide_cover.background.fill.fore_color.rgb = C_PRIMARY
    
    # 封面主标题
    tb_title = slide_cover.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.0))
    p_t = tb_title.text_frame.paragraphs[0]
    p_t.text = topic
    p_t.font.size, p_t.font.bold, p_t.font.color.rgb = Pt(44), True, RGBColor(255, 255, 255)
    
    # 封面副标题
    p_sub = tb_title.text_frame.add_paragraph()
    p_sub.text = f"—— 高中化学优质精品课件 | {grade}"
    p_sub.font.size, p_sub.font.color.rgb = Pt(22), RGBColor(148, 163, 184)
    p_sub.space_before = Pt(20)

    # 2. 正文页（多栏卡片排版）
    for page in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = C_BG
        
        # 页面标题
        tb_head = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.8))
        p_h = tb_head.text_frame.paragraphs[0]
        p_h.text = page.get("title", "教学要点")
        p_h.font.size, p_h.font.bold, p_h.font.color.rgb = Pt(26), True, C_PRIMARY
        
        # 装饰橙色线条
        line = slide.shapes.add_shape(1, Inches(0.8), Inches(1.2), Inches(11.73), Inches(0.03))
        line.fill.solid()
        line.fill.fore_color.rgb, line.line.color.rgb = C_ACCENT, C_ACCENT
        
        cards = page.get("cards", [])
        card_count = max(1, min(len(cards), 3)) # 最多支持 3 栏排版
        
        total_width = 11.73
        gap = 0.3
        card_w = (total_width - (card_count - 1) * gap) / card_count
        card_h = 5.4
        top_y = 1.5
        
        for i, card_info in enumerate(cards[:card_count]):
            left_x = 0.8 + i * (card_w + gap)
            
            card = slide.shapes.add_shape(1, Inches(left_x), Inches(top_y), Inches(card_w), Inches(card_h))
            card.fill.solid()
            card.fill.fore_color.rgb = C_CARD_BG
            card.line.color.rgb = RGBColor(226, 232, 240)
            
            tf = card.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.3)
            
            p_ctitle = tf.paragraphs[0]
            p_ctitle.text = f"▌ {card_info.get('subtitle', '核心要点')}"
            p_ctitle.font.size, p_ctitle.font.bold, p_ctitle.font.color.rgb = Pt(20), True, C_ACCENT
            p_ctitle.space_after = Pt(14)
            
            for pt in card_info.get("items", []):
                p_item = tf.add_paragraph()
                p_item.text = f"• {pt}"
                p_item.font.size = Pt(15)
                p_item.font.color.rgb = C_TEXT_MAIN
                p_item.space_after = Pt(10)
                
    ppt_out = io.BytesIO()
    prs.save(ppt_out)
    ppt_out.seek(0)
    return ppt_out

# --- Word 生成引擎 ---
def build_doc(topic, grade, slides_data, quiz_data):
    doc = Document()
    doc.add_heading(f'《{topic}》教学设计教案', 0)
    doc.add_paragraph(f'适用年级：{grade}')
    
    doc.add_heading('一、教学大纲与课件结构', level=1)
    for s in slides_data:
        doc.add_heading(s['title'], level=2)
        for c in s.get('cards', []):
            doc.add_paragraph(f"【{c.get('subtitle')}】", style='List Bullet')
            for item in c.get('items', []):
                doc.add_paragraph(f"  - {item}")
            
    doc.add_heading('二、高考级随堂检测题', level=1)
    for idx, q in enumerate(quiz_data, 1):
        doc.add_paragraph(f"{idx}. {q['question']}")
        for opt in q.get('options', []):
            doc.add_paragraph(opt)
        p_ans = doc.add_paragraph()
        p_ans.add_run(f"【答案】{q['answer']}\n【解析】{q['explanation']}").bold = True
        doc.add_paragraph('')
        
    doc_out = io.BytesIO()
    doc.save(doc_out)
    doc_out.seek(0)
    return doc_out

# --- 按钮与执行 ---
if st.button("🚀 一键生成高颜值全套教研资源（PPT + 教案 + 题库）", type="primary", use_container_width=True):
    if not api_key:
        st.error("请先在左侧输入你的 API Key！")
    else:
        with st.spinner("🤖 正在结合您的个人要求生成全套教研方案..."):
            try:
                client = OpenAI(api_key=api_key, base_url=base_url)
                
                prompt = f"""
                你是一位全国顶级的化学特级教师与 PPT 视觉设计专家。
                请为【{grade}】课题【{topic}】设计一份高颜值、深度教学的 PPT 课件结构和随堂检测题。

                【教师自定义补充要求】：
                {custom_requirements if custom_requirements.strip() else "无特殊补充，请按照标准高考大纲深度设计。"}

                请严格按照以下 JSON 格式输出（不要添加任何额外的 markdown 或 ```json 标记）：
                {{
                  "slides": [
                    {{
                      "title": "页面标题",
                      "cards": [
                        {{
                          "subtitle": "卡片小标题",
                          "items": ["详细要点1", "详细要点2"]
                        }}
                      ]
                    }}
                  ],
                  "quiz": [
                    {{
                      "question": "题目...",
                      "options": ["A. xx", "B. xx", "C. xx", "D. xx"],
                      "answer": "A",
                      "explanation": "解析..."
                    }}
                  ]
                }}

                要求：
                1. 必须优先融入【教师自定义补充要求】中的内容。
                2. slides 生成 4 页正文，每页包含 2~3 个并列卡片（cards），内容要有深度和考点提炼。
                3. quiz 生成 2 道带详细高考解析的高质量选择题。
                """
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "你是一个只输出结构化 JSON 数据的化学教学设计专家。"},
                        {"role": "user", "content": prompt}
                    ]
                )
                
                ai_content = response.choices[0].message.content.strip()
                if ai_content.startswith("```"):
                    ai_content = ai_content.split("\n", 1)[1].rsplit("\n", 1)[0]
                
                res_data = json.loads(ai_content)
                slides_data = res_data["slides"]
                quiz_data = res_data["quiz"]
                
                ppt_file = build_ppt(topic, grade, slides_data)
                doc_file = build_doc(topic, grade, slides_data, quiz_data)
                
                st.success("🎉 高颜值全套教研资源生成完毕！")
                
                tab1, tab2, tab3 = st.tabs(["📊 精美 PPT 结构预览", "✍️ 随堂检测题", "📥 一键导出下载"])
                
                with tab1:
                    st.info("💡 PPT 包含深色高质感封面页 + 正文多栏卡片排版")
                    for s in slides_data:
                        with st.expander(f"📖 {s['title']}", expanded=True):
                            cols = st.columns(len(s.get("cards", [])))
                            for idx, card in enumerate(s.get("cards", [])):
                                with cols[idx]:
                                    st.subheader(card.get("subtitle"))
                                    for item in card.get("items", []):
                                        st.write(f"- {item}")
                                
                with tab2:
                    for idx, q in enumerate(quiz_data, 1):
                        st.markdown(f"**{idx}. {q['question']}**")
                        for opt in q.get("options", []):
                            st.write(opt)
                        st.info(f"**答案**：{q['answer']} \n\n **解析**：{q['explanation']}")
                        st.divider()
                        
                with tab3:
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        st.download_button(
                            label="📥 下载精美 PPT 课件 (.pptx)",
                            data=ppt_file,
                            file_name=f"{grade}_{topic}_高颜值课件.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                    with col_dl2:
                        st.download_button(
                            label="📄 下载 Word 教案与试题 (.docx)",
                            data=doc_file,
                            file_name=f"{grade}_{topic}_教案试题.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"⚠️ 发生错误：{e}")
