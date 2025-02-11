import streamlit as st

input_num = st.number_input('Input a number', value=0)

result = input_num ** 2
st.write('Result: ', result)
