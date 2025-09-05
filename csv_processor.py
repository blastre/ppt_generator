import pandas as pd
import re
from model import ai

def load_csv(filepath: str) -> pd.DataFrame:
    """Load CSV file into pandas DataFrame"""
    return pd.read_csv(filepath)

def convert_query_to_pandas(question: str, df: pd.DataFrame) -> str:
    """Convert natural language question to pandas operations"""
    columns = list(df.columns)
    sample_data = df.head(3).to_string()
    
    prompt = f"""
Convert this question to pandas DataFrame operations. The DataFrame is called 'df'.
CSV Columns: {columns}
Sample data:
{sample_data}

Question: {question}

Return only the Python pandas code that will answer this question. Use df as the DataFrame name.
Example: df.groupby('column').sum() or df[df['column'] > 100].mean()
"""
    
    return ai(prompt)

def execute_pandas_query(df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Execute pandas query on DataFrame"""
    try:
        # Clean the query - remove any markdown formatting
        query = re.sub(r'```.*?\n|```', '', query).strip()
        
        # Execute the query
        result = eval(query)
        
        # Convert to DataFrame if it's not already
        if not isinstance(result, pd.DataFrame):
            if hasattr(result, 'to_frame'):
                result = result.to_frame()
            else:
                result = pd.DataFrame([result])
        
        return result
    except Exception as e:
        print(f"Query execution error: {e}")
        return df.head(10)  # Return sample data as fallback

def analyze_data(df: pd.DataFrame, question: str) -> dict:
    """Analyze CSV data based on question"""
    
    # Convert question to pandas query
    pandas_query = convert_query_to_pandas(question, df)
    print(f"Generated query: {pandas_query}")
    
    # Execute query
    result_df = execute_pandas_query(df, pandas_query)
    
    # Generate summary
    summary_prompt = f"""
Based on this data analysis result, provide a brief summary:

Question: {question}
Result:
{result_df.to_string()}

Provide a 2-3 sentence summary of what this data shows.
"""
    
    summary = ai(summary_prompt)
    
    return {
        'query': pandas_query,
        'result': result_df,
        'summary': summary,
        'original_question': question
    }