import base64
import re
import shutil
import tempfile
import uuid
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from subprocess import PIPE, run

import streamlit as st


st.set_page_config(page_title="Office to PDF Converter",
                    page_icon='📄',
                    layout='wide',
                    initial_sidebar_state='expanded')

# apply custom css if needed
with open(Path('utils/style.css')) as css:
    st.markdown(f'<style>{css.read()}</style>', unsafe_allow_html=True)


# TODO: this is untested yet
@st.cache_resource(ttl=60*60*24)
def cleanup_tempdir() -> None:
    '''Cleanup temp dir for all user sessions.
    Filters the temp dir for uuid4 subdirs.
    Deletes them if they exist and are older than 1 day.
    '''
    # delete temp dir subfolder if older than 1 day
    deleteTime = datetime.now() - timedelta(days=1)
    # compile regex for uuid4
    uuid4_regex = r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}'
    uuid4_regex = re.compile(uuid4_regex)
    tempfiledir = Path(tempfile.gettempdir())
    if tempfiledir.exists():
        subdirs = [x for x in tempfiledir.iterdir() if x.is_dir()]
        subdirs_match = [x for x in subdirs if uuid4_regex.match(x.name)]
        for subdir in subdirs_match:
            itemTime = datetime.fromtimestamp(subdir.stat().st_mtime)
            if itemTime < deleteTime:
                shutil.rmtree(subdir)


# TODO: this is untested yet
@st.cache_data(show_spinner=False)
def cleanup_session_tempdir() -> None:
    '''Cleanup temp dir for user session.
    Deletes the temp dir if it exists.
    '''
    if 'tempfiledir' in st.session_state:
        tempfiledir = st.session_state['tempfiledir']
        if tempfiledir.exists():
            shutil.rmtree(tempfiledir)


@st.cache_data(show_spinner=False)
def make_tempdir() -> Path:
    '''Make temp dir for each user session and return path to it
    returns: Path to temp dir
    '''
    if 'tempfiledir' not in st.session_state:
        tempfiledir = Path(tempfile.gettempdir())
        tempfiledir = tempfiledir.joinpath(f"{uuid.uuid4()}")   # make unique subdir
        tempfiledir.mkdir(parents=True, exist_ok=True)  # make dir if not exists
        st.session_state['tempfiledir'] = tempfiledir
    return st.session_state['tempfiledir']


@st.cache_data(show_spinner=False)
def store_file_in_tempdir(tmpdirname: Path, uploaded_file: BytesIO) -> Path:
    '''Store file in temp dir and return path to it
    params: tmpdirname: Path to temp dir
            uploaded_file: BytesIO object
    returns: Path to stored file
    '''
    # store file in temp dir
    tmpfile = tmpdirname.joinpath(uploaded_file.name)
    with open(tmpfile, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    return tmpfile


@st.cache_data(show_spinner=False)
def convert_doc_to_pdf_native(doc_file: Path, output_dir: Path=Path("."), timeout: int=60):
    """Converts a doc file to pdf using libreoffice without msoffice2pdf.
    Calls libroeoffice (soffice) directly in headless mode.
    params: doc_file: Path to doc file
            output_dir: Path to output dir
            timeout: timeout for subprocess in seconds
    returns: (output, exception)
            output: Path to converted file
            exception: Exception if conversion failed
    """
    exception = None
    output = None
    try:
        process = run(['soffice', '--headless', '--convert-to',
            'pdf:writer_pdf_Export', '--outdir', output_dir.resolve(), doc_file.resolve()],
            stdout=PIPE, stderr=PIPE,
            timeout=timeout, check=True)
        stdout = process.stdout.decode("utf-8")
        re_filename = re.search('-> (.*?) using filter', stdout)
        output = Path(re_filename[1]).resolve()
    except Exception as e:
        exception = e
    return (output, exception)


@st.cache_data(show_spinner=False)
def get_pdf_bytes(file_path: Path):
    with open(file_path, "rb") as f:
        pdf_bytes = f.read()
    return pdf_bytes


@st.cache_data(show_spinner=False)
def get_base64_encoded_bytes(file_bytes) -> str:
    base64_encoded = base64.b64encode(file_bytes).decode('utf-8')
    return base64_encoded


@st.cache_data(show_spinner=False)
def show_pdf(file_path):
    with open(file_path,"rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="800" height="800" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def show_pdf_base64(base64_pdf):
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="1000px" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def get_versions() -> str:
    result = run(["soffice", "--version"], capture_output=True, text=True)
    libreoffice_version = result.stdout.strip()
    versions = f'''
    - Streamlit:&nbsp;&nbsp;&nbsp;&nbsp;`{st.__version__}`
    - LibreOffice:&nbsp;&nbsp;&nbsp;&nbsp;`{libreoffice_version}`
    '''
    return versions


def show_sidebar():
    with st.sidebar:
        st.header('About')
        st.markdown('''This app can convert **Microsoft Word** or **LibreOffice Writer** Documents to PDF.''')
        st.markdown('''Supported input file formats are:''')
        st.markdown('''`docx`, `doc`, `odt`, `rtf`''')
        st.markdown('''---''')
        st.subheader('Disclaimer')
        st.markdown('''The conversion is done using a headless version of **LibreOffice**.
            As we all know, there are some compatibility issues between LibreOffice and Microsoft Office.
            Therefore, the conversion might not be perfect and may fail in some cases.
            Especially if you convert a Microsoft Office document and/or the document is more complex.''')
        st.markdown('''---''')
        st.subheader('Versions')
        st.markdown(get_versions(), unsafe_allow_html=True)
        st.markdown('''---''')
        st.subheader('GitHub')
        st.markdown('''<https://github.com/Franky1/Streamlit-docx-converter>''')


if __name__ == "__main__":
    tmpdirname = make_tempdir()  # make temp dir for each user session
    show_sidebar()
    st.title('Office to PDF Converter 📄')
    st.markdown('''---''')
    # add streamlit 2 column layout
    col1, col2, col3 = st.columns([2,2,1], gap='large')
    pdf_file = None
    pdf_bytes = None
    with col1:
        st.subheader('Upload a Microsoft Word or LibreOffice Writer file')
        uploaded_file = st.file_uploader(label='Upload a Microsoft Word or LibreOffice Writer file',
                                    type=['doc', 'docx', 'odt', 'rtf'])
        if uploaded_file is not None:
            # store file in temp dir
            tmpfile = store_file_in_tempdir(tmpdirname, uploaded_file)
            # convert file to pdf
            with st.spinner('Converting file...'):
                pdf_file, exception = convert_doc_to_pdf_native(doc_file=tmpfile, output_dir=tmpdirname)
            if pdf_file is None:
                st.error('Conversion failed.')
                st.stop()
            elif exception is not None:
                st.error('Exception occured during conversion.')
                st.error(exception)
                st.stop()
    with col2:
        st.subheader('Download converted PDF file')
        if pdf_file is not None:
            # show result
            st.info(f"Converted file: {pdf_file.name}")
            pdf_bytes = get_pdf_bytes(pdf_file)
            st.download_button(label="Download PDF",
                data=pdf_bytes,
                file_name=pdf_file.name,
                mime='application/octet-stream')
        else:
            st.stop()
    with col3:
        if pdf_file is not None:
            st.image('pdf.png', width=150)

    if pdf_bytes is not None:
        st.markdown('''---''')
        st.subheader('Preview of converted PDF file')
        pdf_bytes_base64 = get_base64_encoded_bytes(pdf_bytes)
        # show pdf in iframe already base64 encoded
        show_pdf_base64(pdf_bytes_base64)
