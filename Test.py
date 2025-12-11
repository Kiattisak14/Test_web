import streamlit as st
import pandas    as pd
import numpy     as np

# 1. Title
st.title('🚌 EV Fleet Guardian: Hello World!')
st.write('ยินดีต้อนรับสู่ระบบทดสอบ Shadow Mode')

# 2. Interactive Button
if st.button('ตรวจสอบสถานะ Server'):
    st.success('✅ Server is Online and Ready!')

# 3. Simple Chart (ลองสร้างกราฟจำลอง)
st.subheader('ตัวอย่างกราฟ Battery Health')
chart_data = pd.DataFrame(
    np.random.randn(20, 3),
    columns=['Bus A', 'Bus B', 'Bus C'])

st.line_chart(chart_data)

# 4. Input Box
user_input = st.text_input("ลองพิมพ์อะไรสักอย่าง:", "เช่น ทดสอบระบบ")
st.write(f"คุณพิมพ์ว่า: {user_input}")