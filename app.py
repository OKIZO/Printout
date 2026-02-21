import streamlit as st
import json
import io
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="MedConcept PPTX生成システム", layout="wide")

# --- 補助関数：図形やセル内のテキストをフォント維持で置換（分割対策版） ---
def replace_text_in_shape(item, replacements):
    if not hasattr(item, "text_frame") or item.text_frame is None:
        return
    for paragraph in item.text_frame.paragraphs:
        # パワポ特有の「文字分割」対策：段落内の文字を一度すべて合体させる
        p_text = "".join(run.text for run in paragraph.runs)
        
        replaced_any = False
        for old_text, new_text in replacements.items():
            if old_text in p_text:
                p_text = p_text.replace(old_text, str(new_text))
                replaced_any = True
                
        if replaced_any:
            # 置換があった場合、最初のブロックに合体させたテキストを入れ、残りのブロックを空にする
            # これによりフォントや文字色（最初の文字のスタイル）が全体に維持されます
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
    # テンプレートの読み込み
    prs = Presentation("template.pptx")

    # 1. テキストの置換マッピングを作成
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

    # 変化タイプ（最大4つ）のマッピング
    cb = json_data.get("changeTypesBefore", [])
    ca = json_data.get("changeTypesAfter", [])
    
    for i in range(4):
        replacements[f"{{{{cb{i+1}}}}}"] = cb[i] if i < len(cb) else ""
        replacements[f"{{{{ca{i+1}}}}}"] = ca[i] if i < len(ca) else ""

    # 2. 全スライドのテキスト置換と不要図形の削除
    for slide in prs.slides:
        shapes_to_delete = []
        
        # グループ化された図形も再帰的にチェックする内部関数
        def process_shapes(shapes):
            for shape in shapes:
                if shape.shape_type == 6: # グループ図形
                    process_shapes(shape.shapes)
                elif hasattr(shape, "text_frame") and shape.text_frame is not None:
                    # 分割対策：段落の文字を合体させてから判定
                    delete_flag = False
                    for paragraph in shape.text_frame.paragraphs:
                        p_text = "".join(run.text for run in paragraph.runs)
                        if "{{cb4}}" in p_text and len(cb) < 4:
                            delete_flag = True
                        if "{{ca4}}" in p_text and len(ca) < 4:
                            delete_flag = True
                    
                    if delete_flag:
                        shapes_to_delete.append(shape)
                    else:
                        replace_text_in_shape(shape, replacements)
                        
                elif shape.has_table: # テーブル内のテキスト置換
                    for row in shape.table.rows:
                        for cell in row.cells:
                            replace_text_in_shape(cell, replacements)

        process_shapes(slide.shapes)

        # マークした図形を削除
        for shape in shapes_to_delete:
            delete_shape(shape)

    # 3. 画像のレイアウト配置（スライド6〜10 / インデックス5〜9）
    slide_indices = {"A案": 5, "B案": 6, "C案": 7, "D案": 8, "E案": 9}
    
    # グリッド計算用の設定（16:9スライド基準）
    margin_x, margin_y = Inches(0.5), Inches(1.5)
    cell_w, cell_h = Inches(3.0), Inches(2.0)
    cols = 3

    for plan_name, images in uploaded_images.items():
        if plan_name in slide_indices and len(prs.slides) > slide_indices[plan_name]:
            slide = prs.slides[slide_indices[plan_name]]
            
            for idx, img_file in enumerate(images[:6]): # 最大6枚まで
                row = idx // cols
                col = idx % cols
                x = margin_x + (col * cell_w)
                y = margin_y + (row * cell_h)
                
                img_stream = io.BytesIO(img_file.read())
                try:
                    slide.shapes.add_picture(img_stream, x, y, width=cell_w - Inches(0.2))
                except Exception as e:
                    st.warning(f"{plan_name}の画像挿入に失敗しました: {e}")

    # 4. メモリ上に保存して出力
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- UI構築 ---
st.title("MedConcept - 企画書PPTX生成システム")

# タブを作成して画面を分ける
tab1, tab2 = st.tabs(["🖼️ STEP 7: 画像アップロード", "📝 STEP 8: テキスト入力＆出力"])

# ===== タブ1: 画像アップロード =====
with tab1:
    st.header("STEP 7: 画像アップロード")
    st.markdown("各デザイン案の画像をアップロードしてください。（各案5〜6枚推奨）")
    
    uploaded_images = {}
    plans = ["A案", "B案", "C案", "D案", "E案"]
    
    ui_cols = st.columns(5)
    for i, plan in enumerate(plans):
        with ui_cols[i]:
            st.subheader(plan)
            uploaded_images[plan] = st.file_uploader(f"{plan}の画像", accept_multiple_files=True, type=["png", "jpg", "jpeg"], key=plan)

# ===== タブ2: JSON入力＆パワポ生成 =====
with tab2:
    st.header("STEP 8: JSONデータ入力 ＆ 企画書生成")
    st.markdown("HTMLアプリで生成されたJSONデータを貼り付け、「企画書を作成」ボタンを押してください。")
    
    json_text = st.text_area("JSONデータを貼り付け", height=300)

    if st.button("📊 企画書を作成", type="primary", use_container_width=True):
        if not json_text.strip():
            st.error("エラー: JSONデータを入力してください。")
        else:
            try:
                json_data = json.loads(json_text)
                with st.spinner("PowerPointを生成中..."):
                    ppt_stream = generate_pptx(json_data, uploaded_images)
                    
                st.success("🎉 PowerPointの生成が完了しました！")
                st.download_button(
                    label="📥 proposal.pptx をダウンロード",
                    data=ppt_stream,
                    file_name=f"proposal_{json_data.get('itemName', 'untitled')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
                
            except json.JSONDecodeError:
                st.error("エラー: JSONのフォーマットが正しくありません。コピー忘れや余分な文字がないか確認してください。")
            except Exception as e:
                st.error(f"予期せぬエラーが発生しました: {e}")
