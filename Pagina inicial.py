
import streamlit as st
import base64


main_bg2 = "realce3.png"

def set_bg_hack(main_bg):
    '''
    A function to unpack an image from root folder and set as bg.
 
    Returns
    -------
    The background.
    '''
    # set bg name
    main_bg_ext = "png"
        
    st.markdown(
         f"""
         <style>
         .stApp {{
             background: url(data:image/{main_bg_ext};base64,{base64.b64encode(open(main_bg, "rb").read()).decode()});
             background-size: cover
         }}
         </style>
         """,
         unsafe_allow_html=True
     )
    
set_bg_hack(main_bg2)

#streamlit run 'Pagina inicial.py'

# streamlit run 'Pagina inicial.py'

#import os
#os.chdir("C:/Users/anaca_b90wyqk/OneDrive - Universidade Federal da Bahia/Documentos/Projeto_stream")


