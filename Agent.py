from openpyxl import load_workbook
import pandas as pd
from langchain_groq import ChatGroq
from dotenv import load_dotenv
load_dotenv()

llm = ChatGroq(model='llama-3.3-70b-versatile')


def Extract_And_Insight(path: str,
                        instruction: str,
                        header_index: int = 0,
                        columns_index: list[int] = [1, 3, 5],
                        chunk_size: int = 1000):

    # --------------------------
    # Excel Extractor
    # --------------------------
    wb = load_workbook(path)
    ws = wb.active

    data = [list(row) for row in ws.iter_rows(values_only=True)]
    df = pd.DataFrame(data)

    # Set header row
    df.columns = df.iloc[header_index]
    df = df.iloc[header_index+1:]
    df = df.iloc[:, columns_index]

    # --------------------------
    # Chunking
    # --------------------------
    chunks = []
    for i in range(0, len(df), chunk_size):
        chunk = df.iloc[i:i+chunk_size].reset_index(drop=True)
        chunks.append(chunk)

    # --------------------------
    # Prepare final output
    # --------------------------
    final_output = []

    for chunk in chunks:
        # Convert chunk to CSV string without header
        Str_df = chunk.to_csv(index=False, header=False)

        prompt = f"""
You are an Employee Reducer.

TASK RULES:
- You will receive two inputs: Data and Instruction.
- For each row in the Data, perform the action required by the Instruction.
- Add TWO new columns on the LEFT side:
    1. Action # should be only limited like Two or three clear Virdict Adjsut as instructed
    2. Reason # inside it write a clear reason as instructed
- Also add an Index column on the far left.
- Return ONLY the final data as comma-separated values (CSV).
- Do NOT add markdown, tables, borders, pipes, code fences, backticks, or explanations.
- Do NOT output ANY extra text — ONLY the CSV table.
- No comments, no notes, no narration.

CRITICAL CSV RULES:
1. If ANY field contains a comma, you MUST wrap it in double quotes.
2. All rows MUST have the same number of columns.
3. Do NOT use tab characters.
4. Do NOT insert spaces before or after commas.
5. Output EXACTLY valid CSV.

FORMAT REQUIREMENT:
Index,Action,Reason,{','.join(df.columns)}

Data:
{Str_df}

Instruction:
{instruction}
"""
        result = llm.invoke(prompt).content
        final_output.append(result)

    # Prepend a single header row to the combined output
    header_row = "Index,Action,Reason," + ",".join(df.columns)
    return header_row + "\n" + "\n".join(final_output)
