import streamlit as st
import pandas as pd
from openai import AzureOpenAI
import os
import openpyxl
import base64
from io import BytesIO


st.title("Excel Intelligence")
client = AzureOpenAI(
    azure_endpoint=os.getenv("ENDPOINT"),
    api_key=os.getenv("OPENAI_API_KEY"),
    api_version="2024-10-01-preview",
)


def detect_graphs(file_path):
    try:
        wb = openpyxl.load_workbook(file_path)
        for sheet in wb.worksheets:
            if sheet._charts:  # Checking for charts in the worksheet
                return True
    except Exception as e:
        print(f"Error processing file: {e}")
        return False

    return False


def encode_file_to_base64(file):
    return base64.b64encode(file.getvalue()).decode("utf-8")


with st.sidebar:
    uploaded_file = st.file_uploader("Upload Financials CSV", type=["xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        if st.button("View Cleaned Data Table"):
            st.write("### Cleaned Data Table")
            st.dataframe(df)

if "messages" not in st.session_state:
    st.session_state.messages = []

if uploaded_file is not None:
    string_data = df.to_string()
    numerical_summary = df.describe().to_string()

    if "initial_analysis" not in st.session_state:
        analysis_prompt = f"""INSTRUCTION:
        Perform a comprehensive analysis of the provided financial dataset. Summarize key insights, trends, and any potential anomalies.
        Ensure the response is clear and easy to understand.
        
        KEY METRICS:
        {numerical_summary}
        
        INPUT: {string_data}"""

        with st.spinner("Analyzing spreadsheet..."):
            analysis_response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": "You are an expert financial data analyst capable of extracting meaningful insights from spreadsheet data.",
                    },
                    {"role": "user", "content": analysis_prompt},
                ],
                temperature=0.0,
            )

        st.session_state.initial_analysis = analysis_response.choices[
            0
        ].message.content.strip()
        st.session_state.messages.append(
            {"role": "assistant", "content": st.session_state.initial_analysis}
        )

        persona_prompt = f"""
        Who can analyse this spreadsheet content with atmost expertise.
        Describe about their persona, tone and professional qualification if any.
        KEY METRICS:
        {numerical_summary}
        
        INPUT: {string_data}"""

        with st.spinner("Analyzing spreadsheet..."):
            analysis_response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": "You are an expert in analysing data description.",
                    },
                    {"role": "user", "content": persona_prompt},
                ],
                temperature=0.0,
            )

        st.session_state.persona_analysis = analysis_response.choices[
            0
        ].message.content.strip()
        st.session_state.messages.append(
            {"role": "assistant", "content": st.session_state.initial_analysis}
        )

        if detect_graphs(uploaded_file):
            st.write("Graph detected")
            if uploaded_file.name.endswith(".xlsx") and detect_graphs(uploaded_file):
                base64_encoded_sheet = encode_file_to_base64(uploaded_file)

            analysis_prompt = f"""INSTRUCTION:
            Explain only about the graph and visual respresentation in the given spreadsheet image. 
            {f'BASE64_SHEET: {base64_encoded_sheet}' if base64_encoded_sheet else ''} """

            with st.spinner("Analyzing spreadsheet..."):
                analysis_response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {
                            "role": "system",
                            "content": "You are an expert financial data analyst capable of extracting meaningful insights from spreadsheet data.",
                        },
                        {"role": "user", "content": analysis_prompt},
                    ],
                    temperature=0.0,
                )

            st.session_state.initial_analysis = analysis_response.choices[
                0
            ].message.content.strip()
            st.session_state.messages.append(
                {"role": "assistant", "content": st.session_state.initial_analysis}
            )

    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    user_input = st.chat_input("Ask a question about the table:")

    if user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})
        with st.chat_message("user"):
            st.markdown(user_input)

        prompt = f"""INSTRUCTION:
        You are a highly skilled AI assistant specializing in spreadsheet data analysis. Your task is to extract accurate answers from structured table data based on user queries.
        The provided input consists of tabular data formatted as a string, where each cell is represented by a pair consisting of a cell address and its content, separated by a comma (e.g., 'A1,Year').
        Cells are organized in a row-major order and are separated by '|', like 'A1,Year|A2,Profit'. Some cells may have empty values, represented as 'A1, |A2,Profit'.
        
        Your response should:
        - Accurately interpret the data in the table.
        - Provide numerical summaries with **consistent totals, averages, min, max, and standard deviations**.
        - Identify relevant patterns, trends, and numerical insights where applicable.
        - Ensure all calculations reference the precomputed metrics.
        
        KEY METRICS:
        {numerical_summary}
        
        INPUT: {string_data}
        USER QUESTION: {user_input}"""
        print(prompt)
        with st.spinner("Fetching answer..."):
            response_stream = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": f"""You are an advanced data analysis assistant, specialized in interpreting and extracting insights from structured spreadsheet data. 
                        You can accurately understand tabular information, identify key trends, and provide concise, human-readable responses based on the given financial dataset.
                        You possess the following characteristics:
                        {st.session_state.persona_analysis}
                        """,
                    },
                    {"role": "assistant", "content": st.session_state.initial_analysis},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.0,
                stream=True,
            )

        bot_response = ""
        with st.chat_message("assistant"):
            response_placeholder = st.empty()
            for chunk in response_stream:
                if chunk.choices:
                    bot_response += chunk.choices[0].delta.content or ""
                    response_placeholder.markdown(bot_response)

        st.session_state.messages.append({"role": "assistant", "content": bot_response})
