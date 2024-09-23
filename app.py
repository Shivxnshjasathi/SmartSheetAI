import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import json
import plotly.express as px

# Configure the Gemini API (you'll need to replace this with your actual API key)
genai.configure(api_key='AIzaSyCjy1KLDa4BKiJVBC_9po77VEqNl49DT3I')
model = genai.GenerativeModel('gemini-pro')

def clean_data(df):
    # Remove rows where all columns are empty
    df_cleaned = df.dropna(how='all')
    
    # Remove columns where all values are empty
    df_cleaned = df_cleaned.dropna(axis=1, how='all')
    
    # For remaining columns, replace empty strings with NaN
    df_cleaned = df_cleaned.replace(r'^\s*$', pd.NA, regex=True)
    
    # Drop rows with any NaN values
    df_cleaned = df_cleaned.dropna()
    
    return df_cleaned

def analyze_and_modify_excel(df, query):
    # Clean the data
    df_cleaned = clean_data(df)
    
    # Convert the cleaned DataFrame to a string representation
    df_string = df_cleaned.to_string()
    
    prompt = f"""
    You are an AI assistant specialized in analyzing and modifying Excel files based on natural language queries. Your task is to provide accurate and reliable results based on the given data and query.

    DataFrame Description:
    - Columns: {', '.join(df_cleaned.columns)}
    - Data Types: {df_cleaned.dtypes.to_dict()}
    - Number of Rows: {len(df_cleaned)}

    Full Data (all rows):
    {df_string}

    The data has been cleaned by removing rows and columns with all empty values, and removing any rows with partially empty values.

    Please analyze this cleaned data carefully and respond to the following query:
    {query}

    When responding to the query, follow these guidelines:

    1. Data Manipulation Queries:
       a. Provide a clear, step-by-step explanation of the required operations.
       b. Generate Python code to perform the requested operation on the entire dataset.
       c. Describe the expected result, including any changes in the data structure or content.
       d. If possible, provide a small sample of the expected output for verification.

    2. Data Visualization Queries:
       Return a JSON string with the following structure:
       {{
         "chart_type": "bar/line/scatter/pie/histogram/box/heatmap",
         "x_column": "column_name",
         "y_column": "column_name",
         "title": "Chart Title",
         "additional_parameters": {{
           // Any additional parameters specific to the chart type
         }}
       }}

    3. Analysis Queries:
       a. Provide a detailed response based on the entire dataset.
       b. Include relevant statistical measures or summaries when appropriate.
       c. Highlight any notable patterns, outliers, or insights from the data.

    4. Error Handling:
       If the query cannot be executed due to data limitations or inconsistencies, explain the issue and suggest alternative approaches or data modifications that might help.

    5. Assumptions:
       Clearly state any assumptions you make when interpreting the query or analyzing the data.

    6. Verification:
       Suggest ways for the user to verify the results or provide sample output for complex operations.

    Please ensure your response is accurate, comprehensive, and directly addresses the user's query. If any part of the query is ambiguous, ask for clarification before proceeding with the analysis or manipulation.
    """
    
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        return None

def execute_modification(df, modification_code):
    try:
        # Create a copy of the dataframe to avoid modifying the original
        modified_df = df.copy()
        
        # Execute the modification code
        exec(modification_code, globals(), {'df': modified_df})
        
        return modified_df
    except Exception as e:
        st.error(f"Error executing modification: {str(e)}")
        return None

def main():
    st.set_page_config(page_title="Natural Language Excel Analyzer", layout="wide")
    st.title("Natural Language Excel Analyzer using Gemini API")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df_cleaned = clean_data(df)
            
            st.sidebar.header("File Information")
            st.sidebar.write(f"Original Rows: {df.shape[0]}")
            st.sidebar.write(f"Cleaned Rows: {df_cleaned.shape[0]}")
            st.sidebar.write(f"Original Columns: {df.shape[1]}")
            st.sidebar.write(f"Cleaned Columns: {df_cleaned.shape[1]}")
            st.sidebar.write("Cleaned Column Names:")
            st.sidebar.write(", ".join(df_cleaned.columns))

            tab1, tab2, tab3 = st.tabs(["Data Preview", "Analysis & Modification", "Visualization"])

            with tab1:
                st.header("Cleaned Data Preview")
                st.dataframe(df_cleaned)  # Display cleaned dataset

            with tab2:
                st.header("Excel Analysis & Modification")
                query = st.text_area("Enter your query or request in natural language:")
                if st.button("Process"):
                    if query:
                        with st.spinner("Processing..."):
                            result = analyze_and_modify_excel(df_cleaned, query)
                        if result:
                            st.subheader("Analysis Result:")
                            st.write(result)
                            
                            # Check if the result contains Python code for modification
                            if "```python" in result:
                                code_start = result.index("```python") + 10
                                code_end = result.index("```", code_start)
                                modification_code = result[code_start:code_end].strip()
                                
                                st.subheader("Modification Preview:")
                                st.code(modification_code, language="python")
                                
                                if st.button("Apply Modification"):
                                    modified_df = execute_modification(df_cleaned, modification_code)
                                    if modified_df is not None:
                                        st.subheader("Modified Data Preview:")
                                        st.dataframe(modified_df)  # Display full modified dataset
                                        
                                        # Option to download the modified Excel file
                                        output = io.BytesIO()
                                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                            modified_df.to_excel(writer, index=False, sheet_name='Sheet1')
                                        output.seek(0)
                                        st.download_button(
                                            label="Download modified Excel file",
                                            data=output,
                                            file_name="modified_excel_file.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                    else:
                        st.warning("Please enter a query or request.")

            with tab3:
                st.header("Data Visualization")
                viz_query = st.text_input("Describe the chart you want to create:")
                if st.button("Generate Chart"):
                    if viz_query:
                        with st.spinner("Generating chart..."):
                            chart_data = analyze_and_modify_excel(df_cleaned, f"Create a chart based on this request: {viz_query}. Return only the JSON string for the chart data.")
                        try:
                            chart_info = json.loads(chart_data)
                            if all(key in chart_info for key in ['chart_type', 'x_column', 'y_column', 'title']):
                                if chart_info['chart_type'] == 'bar':
                                    fig = px.bar(df_cleaned, x=chart_info['x_column'], y=chart_info['y_column'], title=chart_info['title'])
                                elif chart_info['chart_type'] == 'line':
                                    fig = px.line(df_cleaned, x=chart_info['x_column'], y=chart_info['y_column'], title=chart_info['title'])
                                elif chart_info['chart_type'] == 'scatter':
                                    fig = px.scatter(df_cleaned, x=chart_info['x_column'], y=chart_info['y_column'], title=chart_info['title'])
                                elif chart_info['chart_type'] == 'pie':
                                    fig = px.pie(df_cleaned, values=chart_info['y_column'], names=chart_info['x_column'], title=chart_info['title'])
                                st.plotly_chart(fig)
                            else:
                                st.error("Invalid chart data received.")
                        except json.JSONDecodeError:
                            st.error("Failed to generate chart. Please try a different query.")
                    else:
                        st.warning("Please enter a visualization query.")

        except Exception as e:
            st.error(f"Error processing the uploaded file: {str(e)}")
            st.info("Please ensure you've uploaded a valid Excel file.")

    else:
        st.info("Please upload an Excel file to begin analysis.")

if __name__ == "__main__":
    main()
