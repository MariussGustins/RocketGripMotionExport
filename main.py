import os
import requests
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from dotenv import load_dotenv
from datetime import datetime

# Load API key
load_dotenv()
API_KEY = os.getenv("MOTION_API_KEY")

HEADERS = {
    "X-API-Key": API_KEY,
    "Content-Type": "application/json"
}

def fetch_data(month, year):
    workspace_url = "https://api.usemotion.com/v1/workspaces"
    workspace_response = requests.get(workspace_url, headers=HEADERS)

    if workspace_response.status_code != 200:
        print("❌ Failed to fetch workspaces:", workspace_response.status_code, workspace_response.text)
        return []

    workspaces = workspace_response.json().get("workspaces", [])
    if not workspaces:
        print("❌ No workspaces found.")
        return []

    all_task_data = []

    for ws in workspaces:
        workspace_id = ws["id"]
        workspace_name = ws.get("name", "Unnamed Workspace")
        print(f"🔄 Fetching data for workspace: {workspace_name}")

        # Get projects
        projects_url = f"https://api.usemotion.com/v1/projects?workspaceId={workspace_id}"
        projects_response = requests.get(projects_url, headers=HEADERS)
        project_map = {}
        if projects_response.status_code == 200:
            project_map = {
                p["id"]: p["name"]
                for p in projects_response.json().get("projects", [])
            }

        # Get tasks (including recurring instances)
        tasks_url = f"https://api.usemotion.com/v1/tasks?includeAllStatuses=true&workspaceId={workspace_id}"
        tasks_response = requests.get(tasks_url, headers=HEADERS)
        if tasks_response.status_code != 200:
            print(f"⚠️ Tasks error: {tasks_response.text}")
            continue

        tasks = tasks_response.json().get("tasks", [])

        for task in tasks:
            completed = task.get("completed", False)
            completed_time_raw = task.get("completedTime")
            last_interacted_raw = task.get("lastInteractedTime")

            if not completed and not last_interacted_raw:
                continue

            # Filter by selected month
            ref_raw = completed_time_raw if completed else last_interacted_raw
            try:
                ref_dt = datetime.fromisoformat(ref_raw.replace("Z", "+00:00"))
                if ref_dt.month != month or ref_dt.year != year:
                    continue
                ref_label = ref_dt.strftime("%Y-%m-%d %H:%M")
            except:
                continue

            status = "Completed" if completed else "In Progress"
            task_type = "Recurring" if task.get("parentRecurringTaskId") else "Regular"

            # Duration logic
            duration_minutes = 0
            if completed:
                duration = task.get("duration")
                if isinstance(duration, int):
                    duration_minutes = duration
                else:
                    continue
            else:
                for chunk in task.get("chunks", []):
                    if chunk.get("completedTime") and isinstance(chunk.get("duration"), int):
                        duration_minutes += chunk["duration"]
                if duration_minutes == 0:
                    continue

            h, m = divmod(duration_minutes, 60)
            duration_str = f"{h}h {m}m" if h else f"{m}m"

            project_data = task.get("project")
            project_id = project_data.get("id") if project_data else None
            project_name = project_map.get(project_id, "No Project")

            assignees = task.get("assignees", [])
            if not assignees:
                assignees = [{"name": "Unassigned"}]

            task_id = task.get("id")

            for a in assignees:
                assignee_name = a.get("name", "Unknown")
                all_task_data.append({
                    "Task ID": task_id,  # optional, can be hidden in final report
                    "Workspace": workspace_name,
                    "Project": project_name,
                    "Task Name": task.get("name"),
                    "Assignees": assignee_name,
                    "Status": status,
                    "Type": task_type,
                    "Last Active": ref_label,
                    "Duration": duration_str,
                    "Duration (min)": duration_minutes
                })

    return all_task_data


def export_to_excel(all_task_data):
    if not all_task_data:
        messagebox.showerror("No Data", "No tasks found.")
        return

    df = pd.DataFrame(all_task_data)

    # Create pivot
    pivot_df = df.pivot_table(
        index=["Workspace", "Project"],
        values="Duration (min)",
        aggfunc="sum"
    )
    pivot_df["Duration (h m)"] = pivot_df["Duration (min)"].apply(
        lambda x: f"{int(x // 60)}h {int(x % 60)}m"
    )


    with pd.ExcelWriter("motion_tasks.xlsx") as writer:
        df.to_excel(writer, sheet_name="Data Dump", index=False)
        pivot_df.to_excel(writer, sheet_name="Pivot", index=True)

    messagebox.showinfo("✅ Done", f"Exported {len(df)} total tasks (incl. completed recurring tasks)")


def run_gui():
    root = tk.Tk()
    root.title("Motion Task Exporter")

    tk.Label(root, text="Select Month and Year", font=("Arial", 14)).pack(pady=10)

    tk.Label(root, text="Month:").pack()
    month_cb = ttk.Combobox(root, values=[f"{i:02}" for i in range(1, 13)], state="readonly")
    month_cb.set("04")
    month_cb.pack()

    tk.Label(root, text="Year:").pack()
    year_cb = ttk.Combobox(root, values=[str(y) for y in range(2020, 2031)], state="readonly")
    year_cb.set("2025")
    year_cb.pack()

    def on_proceed():
        month = int(month_cb.get())
        year = int(year_cb.get())
        data = fetch_data(month, year)
        export_to_excel(data)

    tk.Button(root, text="Proceed", command=on_proceed, font=("Arial", 12)).pack(pady=20)
    root.mainloop()


if __name__ == "__main__":
    run_gui()