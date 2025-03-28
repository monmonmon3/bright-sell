import streamlit as st
import cash_data_import
import invoice_data_import

# レイアウトを centered（デフォルト）に設定
st.set_page_config(layout="centered") 

# 初期状態では何も表示しないようにセッション状態を設定
if 'current_app' not in st.session_state:
    st.session_state['current_app'] = None

# タイトルを左寄せにするためのCSS
st.markdown("""
    <style>
    .stTitle {
        text-align: left;
    }
    </style>
""", unsafe_allow_html=True)

# タイトルの表示
st.title('医療法人社団ブライトデンタルケア 売上データ freee_import')

# サイドバーでアプリ選択用のボタンを表示
with st.sidebar:
    # サイドバーのタイトル
    st.sidebar.title('処理メニュー')
    # サイドバーの説明
    st.sidebar.write('実行する処理を選択してください')
    
    if st.button('窓口売上のインポートデータ作成', key='app1'):
        st.session_state['current_app'] = 'app1'

    if st.button('請求売上のインポートデータ作成', key='app2'):
        st.session_state['current_app'] = 'app2'

# 選択されたアプリを表示
if st.session_state['current_app'] == 'app1':
    cash_data_import.app1()
elif st.session_state['current_app'] == 'app2':
    invoice_data_import.app2()

# インラインでCSSを追加
st.markdown("""
    <style>
    body {
        background-color: #f0f0f0;
        font-family: Arial, sans-serif;
    }

    .stButton button {
        background-color: #000000;  /* ボタンの背景を黒に設定 */
        color: white;               /* ボタンのテキストを白に設定 */
        font-size: 16px;
        border-radius: 5px;
        padding: 10px 20px;
        border: none;              /* ボタンの枠線をなくす */
        transition: transform 0.2s ease;  /* アニメーション効果を追加 */
    }

    /* サイドバーを黒に設定 */
    .css-1d391kg {  /* サイドバーの背景色を変更 */
        background-color: #000000;
        color: white;
    }

    /* サイドバー内のボタンの色を白に設定 */
    .css-1d391kg .stButton>button {
        background-color: #000000;  /* サイドバー内のボタン背景を黒に設定 */
        color: white;               /* テキスト色を白に設定 */
    }

    /* ボタンが押されたときの色の変更を防止 */
    .stButton button:active,
    .stButton button:focus,
    .stButton button:hover {
        background-color: #000000 !important;  /* 押されたときも背景色を黒に設定 */
        color: white !important;               /* テキスト色を白に設定 */
    }

    /* ボタンにカーソルが乗ったときに少し大きくする */
    .stButton button:hover {
        transform: scale(1.1);  /* ボタンを10%大きくする */
    }

    </style>
""", unsafe_allow_html=True)
