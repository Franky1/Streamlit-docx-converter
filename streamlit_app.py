import streamlit as st
from msoffice2pdf import convert


# TODO: set basic page config
st.set_page_config(page_title="Streamlit Template",
                    page_icon='✅',
                    initial_sidebar_state='collapsed')

# TODO: apply custom css if needed
# with open('utils/style.css') as css:
#     st.markdown(f'<style>{css.read()}</style>', unsafe_allow_html=True)


if __name__ == "__main__":
    # TODO: write streamlit app here
    st.title('🔨 Streamlit Template')
    st.markdown("""
        This app is only a template for a new Streamlit project. <br>

        ---
        """, unsafe_allow_html=True)

    st.balloons()
