Below is an in‐depth, multi‐file solution. One file (named, for example, **task_manager.py**) contains all the backend logic to interact with Azure OpenAI and Excel, and a second file (named **app.py**) contains the Streamlit UI that provides two pages: one for visualizing your detailed to‑do list and one for chatting with the assistant.

Before you run the code, be sure to install the required packages:

```bash
pip install streamlit pandas openpyxl requests
```

Also, set your Azure OpenAI configuration values (or replace the placeholders in the code).

---

### File: task_manager.py

This file defines functions for:
- **Excel management:** loading, saving, summarizing, and updating tasks in an Excel file.
- **Azure OpenAI integration:** sending conversation history and receiving responses.
- **JSON extraction:** parsing the assistant’s response when it uses structured JSON (wrapped in triple backticks tagged with `json`).

```python
# task_manager.py
import os
import re
import json
import requests
import pandas as pd
from datetime import datetime

# =====================================================
# Azure OpenAI Configuration
# =====================================================
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT", "https://<your-resource-name>.openai.azure.com")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY", "<your-azure-openai-api-key>")
DEPLOYMENT_NAME = os.getenv("DEPLOYMENT_NAME", "<your-deployment-name>")  # e.g., "gpt-35-turbo"
API_VERSION = "2023-03-15-preview"

# =====================================================
# Excel File Configuration
# =====================================================
EXCEL_FILE = "tasks.xlsx"

# =====================================================
# System Prompt for OpenAI
# =====================================================
SYSTEM_PROMPT = (
    "You are an intelligent task management assistant with access to an Excel-based to-do list. "
    "Your responses can be either structured commands or plain text answers. "
    "When given a task command, respond with a structured JSON object enclosed in triple backticks and tagged as json, following exactly this schema:\n"
    "```json\n"
    "{\n"
    '  "action": "add" | "update" | "complete" | "update_cell",\n'
    '  "id": "Task ID (if applicable, else empty)",\n'
    '  "task": "Task title",\n'
    '  "due_date": "Due date in YYYY-MM-DD format (or empty)",\n'
    '  "assigned_to": "Name of person assigned (or empty)",\n'
    '  "description": "Detailed description (or empty)",\n'
    '  "clarification": "If information is missing, ask for clarification (or empty)"\n'
    "}\n"
    "```\n"
    "If the input is a general inquiry (for example, 'What tasks are due this week?'), answer in plain text using the tasks summary provided as context. "
    "Make sure to refer to the tasks summary context when answering questions about current tasks."
)

# =====================================================
# Excel Management Functions
# =====================================================
def load_tasks(file_path=EXCEL_FILE):
    """Load tasks from an Excel file; create one if it doesn't exist."""
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
    else:
        df = pd.DataFrame(columns=["ID", "Task", "Due Date", "Assigned To", "Description", "Status", "Created Date"])
        df.to_excel(file_path, index=False)
    return df

def save_tasks(df, file_path=EXCEL_FILE):
    """Save tasks DataFrame to Excel."""
    df.to_excel(file_path, index=False)

def get_tasks_summary():
    """Generate a summary of current tasks as a string."""
    df = load_tasks()
    if df.empty:
        return "No tasks available."
    summary_lines = []
    for _, row in df.iterrows():
        summary_lines.append(
            f"ID {int(row['ID'])}: {row['Task']} due on {row['Due Date']} assigned to {row['Assigned To']} (Status: {row['Status']})"
        )
    return "\n".join(summary_lines)

def update_tasks_from_response(task_info):
    """
    Update the Excel tasks based on the structured JSON command.
    Returns a message describing the outcome.
    """
    df = load_tasks()
    action = task_info.get("action", "").lower()

    if action == "add":
        new_id = int(df["ID"].max()) + 1 if not df.empty and pd.notna(df["ID"].max()) else 1
        new_row = {
            "ID": new_id,
            "Task": task_info.get("task", ""),
            "Due Date": task_info.get("due_date", ""),
            "Assigned To": task_info.get("assigned_to", ""),
            "Description": task_info.get("description", ""),
            "Status": "Pending",
            "Created Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        df = df.append(new_row, ignore_index=True)
        result_msg = f"Added new task with ID {new_id}."

    elif action == "update":
        task_id = task_info.get("id")
        if not task_id:
            return "Error: Task ID is required for an update."
        if int(task_id) in df["ID"].values:
            idx = df.index[df["ID"] == int(task_id)][0]
            # Update provided fields if they are not empty
            df.at[idx, "Task"] = task_info.get("task") or df.at[idx, "Task"]
            df.at[idx, "Due Date"] = task_info.get("due_date") or df.at[idx, "Due Date"]
            df.at[idx, "Assigned To"] = task_info.get("assigned_to") or df.at[idx, "Assigned To"]
            df.at[idx, "Description"] = task_info.get("description") or df.at[idx, "Description"]
            result_msg = f"Updated task with ID {task_id}."
        else:
            result_msg = f"Error: Task ID {task_id} not found."

    elif action == "complete":
        task_id = task_info.get("id")
        if not task_id:
            return "Error: Task ID is required to mark a task complete."
        if int(task_id) in df["ID"].values:
            idx = df.index[df["ID"] == int(task_id)][0]
            df.at[idx, "Status"] = "Completed"
            result_msg = f"Marked task ID {task_id} as completed."
        else:
            result_msg = f"Error: Task ID {task_id} not found."

    elif action == "update_cell":
        # For updating a specific cell, the JSON must include "id", "field", and "value"
        task_id = task_info.get("id")
        field = task_info.get("field")
        value = task_info.get("value")
        if not task_id or not field:
            return "Error: Both task ID and field are required for a cell update."
        if int(task_id) in df["ID"].values and field in df.columns:
            idx = df.index[df["ID"] == int(task_id)][0]
            df.at[idx, field] = value
            result_msg = f"Updated task ID {task_id}: set {field} to {value}."
        else:
            result_msg = "Error: Invalid task ID or field name."
    else:
        result_msg = "Error: Action not recognized."

    save_tasks(df)
    return result_msg

# =====================================================
# Azure OpenAI Query Function
# =====================================================
def query_azure_openai(conversation):
    """
    Query the Azure OpenAI Chat API with the provided conversation history.
    Returns the assistant's reply as a string.
    """
    headers = {
        "Content-Type": "application/json",
        "api-key": AZURE_OPENAI_API_KEY,
    }
    url = f"{AZURE_OPENAI_ENDPOINT}/openai/deployments/{DEPLOYMENT_NAME}/chat/completions?api-version={API_VERSION}"
    
    data = {
        "messages": conversation,
        "max_tokens": 400,
        "temperature": 0.2,
    }
    
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        response_data = response.json()
        return response_data["choices"][0]["message"]["content"].strip()
    else:
        return f"Error querying Azure OpenAI: {response.text}"

# =====================================================
# JSON Extraction Function
# =====================================================
def extract_json_block(text):
    """
    Extract the JSON content enclosed in triple backticks with 'json' tag.
    Returns a dict if found and valid, otherwise None.
    """
    pattern = r"```json\s*(\{.*?\})\s*```"
    match = re.search(pattern, text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            return None
    return None
```

---

### File: app.py

This is the Streamlit UI file. It creates two pages:
- **To‑Do List:** Visualizes the current tasks (from the Excel file) and upcoming tasks.
- **Chat Interface:** Displays a conversation history (maintained in session state), lets you send messages, passes a dynamic context (the current tasks summary) to Azure OpenAI, extracts any structured JSON command, and then updates the Excel file accordingly.

```python
# app.py
import streamlit as st
from datetime import timedelta
import pandas as pd
import task_manager

# Initialize conversation history if not already present
if "chat_history" not in st.session_state:
    st.session_state.chat_history = [{"role": "system", "content": task_manager.SYSTEM_PROMPT}]

# ---------------------------
# To-Do List Page
# ---------------------------
def show_todo_list():
    st.title("To-Do List Overview")
    df = task_manager.load_tasks()
    st.subheader("Detailed Task Tracker")
    st.dataframe(df)
    
    # Display upcoming tasks in the next 7 days
    try:
        df['Due Date'] = pd.to_datetime(df['Due Date'], errors='coerce')
        upcoming = df[(df['Due Date'] >= pd.Timestamp.today()) & 
                      (df['Due Date'] <= (pd.Timestamp.today() + timedelta(days=7)))]
        if not upcoming.empty:
            st.subheader("Upcoming Tasks (Next 7 Days)")
            st.dataframe(upcoming)
        else:
            st.info("No upcoming tasks in the next 7 days.")
    except Exception as e:
        st.error("Error processing dates: " + str(e))

# ---------------------------
# Chat Interface Page
# ---------------------------
def show_chat_interface():
    st.title("Chat Interface")
    
    # Display conversation history
    st.subheader("Conversation History")
    for msg in st.session_state.chat_history:
        if msg["role"] == "user":
            st.markdown(f"**You:** {msg['content']}")
        elif msg["role"] == "assistant":
            st.markdown(f"**Assistant:** {msg['content']}")
        else:
            st.markdown(f"**System:** {msg['content']}")
    
    # Input for new message
    user_message = st.text_input("Enter your message:", key="user_message_input")
    if st.button("Send Message") and user_message.strip():
        # Append the user message to the conversation history
        st.session_state.chat_history.append({"role": "user", "content": user_message})
        
        # Build conversation history with dynamic context (current tasks summary)
        tasks_summary = task_manager.get_tasks_summary()
        context_message = {"role": "system", "content": f"Current tasks summary:\n{tasks_summary}"}
        conversation = [st.session_state.chat_history[0], context_message] + st.session_state.chat_history[1:]
        
        with st.spinner("Waiting for Azure OpenAI response..."):
            assistant_reply = task_manager.query_azure_openai(conversation)
        
        # Append the assistant's reply to the history and display it
        st.session_state.chat_history.append({"role": "assistant", "content": assistant_reply})
        st.markdown(f"**Assistant:** {assistant_reply}")
        
        # Attempt to extract a structured JSON command from the reply
        command = task_manager.extract_json_block(assistant_reply)
        if command:
            if command.get("clarification", ""):
                st.info(f"Clarification requested: {command['clarification']}")
            elif command.get("action", ""):
                result = task_manager.update_tasks_from_response(command)
                st.info(result)
    
    if st.button("Reset Conversation"):
        st.session_state.chat_history = [{"role": "system", "content": task_manager.SYSTEM_PROMPT}]
        st.experimental_rerun()

# ---------------------------
# Main App Navigation
# ---------------------------
def main():
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Select Page", ["To-Do List", "Chat Interface"])
    if page == "To-Do List":
        show_todo_list()
    elif page == "Chat Interface":
        show_chat_interface()

if __name__ == "__main__":
    main()
```

---

### How to Run

1. Place both **task_manager.py** and **app.py** in the same directory.  
2. Set your Azure OpenAI endpoint, API key, and deployment name (either as environment variables or by replacing the placeholders).  
3. Run the Streamlit app with:

   ```bash
   streamlit run app.py
   ```

Now you have a robust, in‑depth solution that lets you have a conversation with an AI assistant that “knows” your current Excel‑backed to‑do list, updates it based on structured JSON commands, and shows you detailed task and progress visualizations—all via Streamlit.
