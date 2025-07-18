import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

from langchain_core.prompts import ChatPromptTemplate
from langchain_ollama.llms import OllamaLLM


def summarize_operations_employees(df):
    summary = []
    operations = df['Operation'].unique()
    for op in operations:
        op_df = df[df['Operation'] == op]
        employees = op_df['Employee'].unique()
        total_cost = op_df['Fees_after'].sum()
        min_cost_row = op_df.loc[op_df['Fees_after'].idxmin()]
        summary.append({
            'Operation': op,
            'NumEmployees': len(employees),
            'TotalCost': total_cost,
            'LowestCostEmployee': min_cost_row['Employee'],
            'LowestCost': min_cost_row['Fees_after']
        })
    return pd.DataFrame(summary)

def analysis_prompt(summary_df):
    prompt = (
        "You are an expert operations analyst. "
        "Given the following summary of operations and employee costs, provide insights such as: "
        "- Which operations have the highest and lowest total costs? "
        "- Which employees have the lowest costs per operation? "
        "- Any notable patterns or anomalies in the data?\n\n"
        "Summary Table:\n"
        f"{summary_df.to_string(index=False)}\n\n"
        "Please provide a concise analysis."
    )
    return prompt

if __name__ == "__main__":
    import glob
    import os

    # Directory containing audit chunk files
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    CHUNKS_DIR = os.path.abspath(os.path.join(BASE_DIR, "../chunks/target_chunks"))

    # Find all Excel files in the chunks directory
    chunk_files = glob.glob(os.path.join(CHUNKS_DIR, "*.xlsx"))
    print(f"Found {len(chunk_files)} chunk files:")
    for f in chunk_files:
        print(f"  - {os.path.basename(f)}")

    # Load each chunk into a DataFrame
    chunk_dfs = []
    for file in chunk_files:
        try:
            df = pd.read_excel(file)
            chunk_dfs.append(df)
        except Exception as e:
            print(f"Failed to load {file}: {e}")

    print(f"Loaded {len(chunk_dfs)} chunk DataFrames.")

    # Combine all chunk DataFrames into one
    if chunk_dfs:
        all_chunks_df = pd.concat(chunk_dfs, ignore_index=True)
    else:
        print("No chunk data loaded. Exiting.")
        exit(1)

    # summary_df = summarize_operations_employees(all_chunks_df)

    # Create the analysis prompt using the summary
    # prompt_text = analysis_prompt(summary_df)

    # summary_df_str = summary_df.to_string(index=False)
    # Use the language model to generate insights
    # template = """
    #     ### Context
    #     You are given data on operations and employee fees from an organizational setting. The dataset includes anonymized information such as operation names, employee identities, fees related to employee participation in operations, and total costs associated with each operation.

    #     ### Task
    #     Your task is to analyze the provided summary of operational costs and employee contributions. The goal is to extract key insights about financial distribution and personnel involvement. 

    #     ### Insights Required
    #     1. **Financial Analysis**:
    #     - Identify the operations with the highest and lowest total costs.
    #     - Discuss notable differences in cost distribution across operations.

    #     2. **Personnel Impact**:
    #     - Highlight employees with the lowest costs per operation.
    #     - Describe potential reasons for cost discrepancies.

    #     3. **Patterns and Anomalies**:
    #     - Pinpoint any unusual patterns in operations or costs that could indicate inefficiencies or areas for improvement.

    #     4. **Recommendations**:
    #     - Suggest steps to optimize expenses and improve employee participation efficiency.

    #     ### Analysis Output
    #     Based on the above context and data summary, provide a comprehensive analysis with insights and recommendations.
    #     """
    # prompt = ChatPromptTemplate.from_template(template)

    # model = OllamaLLM(model="mistral:latest", temperature=0.3)

    # chain = prompt | model

    # result = chain.invoke({"question": prompt_text})

    # print("Insights about Operations and Fees:")
    # print(result)

    # Compute average spent for each operation per month
    avg_spent_per_month = (
        all_chunks_df[all_chunks_df['Fees_after'] != -1]
        .groupby(['Operation', 'Month'])
        .agg(AverageSpent=('Fees_after', 'mean'))
        .reset_index()
    )

    # Plot the average spent for each operation per month
    plt.figure(figsize=(12, 7))
    sns.lineplot(data=avg_spent_per_month, x='Month', y='AverageSpent', hue='Operation', marker='o')
    plt.title("Average Spent per Operation per Month")
    plt.xlabel("Month")
    plt.ylabel("Average Fees Spent (€)")
    plt.xticks(rotation=45)
    plt.legend(title='Operation', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.show()