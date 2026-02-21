import streamlit as st
import json
import io
from pptx import Presentation
from pptx.util import Inches

# レイアウトを横幅いっぱいに使う設定（2カラムに最適化）
st.set_page_config(page_title="PPTX生成システム", layout="wide")

# --- 補助関数：図形やセル内のテキストをフォント維持で置換（分割対策版） ---
def replace_text_in_shape(item, replacements):
    if not hasattr(item, "text_frame") or item.text_frame is None:
        return
    for paragraph in item.text_frame.paragraphs:
        p_text = "".join(run.text for run in paragraph.runs)
        
        replaced_any = False
        for old_text, new_text in replacements.items():
            if old_text in p_text:
                p_text = p_text.replace(old_text, str(new_text))
                replaced_any = True
                
        if replaced_any:
            if len(paragraph.runs) > 0:
                paragraph.runs[0].text = p_text
                for i in range(1, len(paragraph.runs)):
                    paragraph.runs[i].text = ""

# --- 補助関数：不要な図形を完全に削除 ---
def delete_shape(shape):
    try:
        sp_tree = shape.element.getparent()
        sp_tree.remove(shape.element)
    except:
        pass

# --- メイン処理関数 ---
def generate_pptx(json_data, uploaded_images):
    prs = Presentation("template.pptx")

    brand_info = f"カラー：{json_data.get('brandColors', '')}\nブランドイメージ：{'、'.join(json_data.get('brandImages', []))}"
    
    replacements = {
        "{{productName}}": json_data.get("productName", ""),
        "{{itemName}}": json_data.get("itemName", ""),
        "{{spec}}": json_data.get("spec", ""),
        "{{target}}": json_data.get("target", ""),
        "{{scene}}": json_data.get("scene", ""),
        "{{objectiveA}}": json_data.get("objectiveA", ""),
        "{{objectiveB}}": json_data.get("objectiveB", ""),
        "{{before}}": json_data.get("before", ""),
        "{{after}}": json_data.get("after", ""),
        "{{concept}}": json_data.get("concept", ""),
        "{{brandInfo}}": brand_info,
        "{{designExterior}}": "、".join(json_data.get("designExterior", [])),
        "{{functional}}": "、".join(json_data.get("functional", [])),
        "{{toneManner}}": "\n".join(json_data.get("toneManner", [])),
    }

    cb = json_data.get("changeTypesBefore", [])
    ca = json_data.get("changeTypesAfter", [])
    
    for i in range(4):
        replacements[f"{{{{cb{i+1}}}}}"] = cb[i] if i < len(cb) else ""
        replacements[f"{{{{ca{i+1}}}}}"] = ca[i] if i < len(ca) else ""

    for slide in prs.slides:
        def process_shapes(shapes):
            for shape in shapes:
                if shape.shape_type == 6:
                    process_shapes(shape.shapes)
                elif hasattr(shape, "text_frame") and shape.text_frame is not None:
                    replace_text_in_shape(shape, replacements)
                elif shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            replace_text_in_shape(cell, replacements)
        process_shapes(slide.shapes)

    slide_indices = {"A案": 5, "B案": 6, "C案": 7, "D案": 8, "E案": 9}
    margin_x, margin_y = Inches(0.5), Inches(1.5)
    cell_w, cell_h = Inches(3.0), Inches(2.0)
    cols = 3

    for plan_name, images in uploaded_images.items():
        if plan_name in slide_indices and len(prs.slides) > slide_indices[plan_name]:
            slide = prs.slides[slide_indices[plan_name]]
            
            for idx, img_file in enumerate(images[:6]):
                row = idx // cols
                col = idx % cols
                x = margin_x + (col * cell_w)
                y = margin_y + (row * cell_h)
                
                img_stream = io.BytesIO(img_file.read())
                try:
                    slide.shapes.add_picture(img_stream, x, y, width=cell_w - Inches(0.2))
                except Exception as e:
                    st.warning(f"{plan_name}の画像挿入に失敗しました: {e}")

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- UI構築（左右2カラムレイアウト） ---

# 余白を詰めるためのCSS
st.markdown("""
    <style>
        .block-container { padding-top: 1.5rem; padding-bottom: 1.5rem; }
        h1 { font-size: 1.8rem !important; margin-bottom: 1rem !important; }
        h2 { font-size: 1.3rem !important; margin-bottom: 0.5rem !important;}
        .stMarkdown p { font-size: 0.9rem; margin-bottom: 0.5rem !important;}
    </style>
""", unsafe_allow_html=True)

st.title("PPTX生成システム")

# 画面を左右に2分割（間隔を少し広めに取る）
col1, col2 = st.columns(2, gap="large")

# ===== 左カラム：画像アップロード =====
with col1:
    st.header("🖼️ 画像アップロード")
    st.markdown("各案の画像（5〜6枚推奨）をアップロードしてください。")

    uploaded_images = {}
    plans = ["A案", "B案", "C案", "D案", "E案"]

    for plan in plans:
        with st.expander(f"📁 {plan} の画像を選択"):
            uploaded_images[plan] = st.file_uploader(
                f"{plan}の画像", 
                accept_multiple_files=True, 
                type=["png", "jpg", "jpeg"], 
                key=plan,
                label_visibility="collapsed"
            )

# ===== 右カラム：JSON入力＆パワポ生成 =====
with col2:
    st.header("📝 企画書生成")
    st.markdown("左側のアプリからコピーしたJSONデータを貼り付けます。")

    # テキストエリアの高さを、左のメニュー群と合うように少し高め（280）に設定
    json_text = st.text_area("JSONデータを貼り付け", height=280, label_visibility="collapsed", placeholder="ここにJSONデータを貼り付けてください")

    if st.button("📊 企画書パワーポイントを作成", type="primary", use_container_width=True):
        if not json_text.strip():
            st.error("エラー: JSONデータを入力してください。")
        else:
            try:
                json_data = json.loads(json_text)
                with st.spinner("PowerPointを生成中..."):
                    ppt_stream = generate_pptx(json_data, uploaded_images)
                    
                st.success("🎉 PowerPointの生成が完了しました！")
                st.download_button(
                    label="📥 企画書(.pptx) をダウンロード",
                    data=ppt_stream,
                    file_name=f"proposal_{json_data.get('itemName', 'untitled')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
                
            except json.JSONDecodeError:
                st.error("エラー: JSONのフォーマットが正しくありません。")
            except Exception as e:
                st.error(f"予期せぬエラーが発生しました: {e}")
