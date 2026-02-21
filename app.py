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
