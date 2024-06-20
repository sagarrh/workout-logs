import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import matplotlib.pyplot as plt
import os
import requests
import json
from datetime import datetime

# File name for storing data
EXCEL_FILE = 'gym_progress.xlsx'

# Initialize Excel file
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append(['Date', 'Exercise', 'Weight', 'Sets', 'Reps', 'Distance', 'Duration'])
    wb.save(EXCEL_FILE)

# Save data to Excel file
def save_data(date, exercise, weight, sets, reps, distance, duration):
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([date, exercise, weight, sets, reps, distance, duration])
    wb.save(EXCEL_FILE)
    messagebox.showinfo("Data Saved", "Your workout data has been saved.")

# Generate graph of progress
def generate_graph():
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    exercises = {}
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        date, exercise, weight, sets, reps, distance, duration = row
        if exercise not in exercises:
            exercises[exercise] = []
        exercises[exercise].append((date, weight, distance, duration))
    
    for exercise, data in exercises.items():
        dates, weights, distances, durations = zip(*data)
        if exercise in ['Walking', 'Cycling']:
            plt.plot(dates, distances, marker='o', label=f'{exercise} Distance')
        else:
            plt.plot(dates, weights, marker='o', label=f'{exercise} Weight')
    
    plt.xlabel('Date')
    plt.ylabel('Progress')
    plt.title('Workout Progress Over Time')
    plt.legend()
    plt.show()

# Fetch data from Runkeeper API
def fetch_runkeeper_data(token):
    headers = {'Authorization': f'Bearer {token}'}
    activities = requests.get('https://api.runkeeper.com/fitnessActivities', headers=headers).json()
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    for activity in activities['items']:
        date = datetime.strptime(activity['start_time'], '%a, %d %b %Y %H:%M:%S').strftime('%Y-%m-%d')
        exercise = activity['type']
        distance = activity['total_distance']
        duration = activity['duration']
        ws.append([date, exercise, None, None, None, distance, duration])
    
    wb.save(EXCEL_FILE)
    messagebox.showinfo("Data Downloaded", "Your Runkeeper data has been downloaded.")

# GUI
root = tk.Tk()
root.title("Gym Progress Tracker")

# Date
tk.Label(root, text="Date (YYYY-MM-DD)").grid(row=0, column=0)
date_entry = tk.Entry(root)
date_entry.grid(row=0, column=1)

# Exercise
tk.Label(root, text="Exercise").grid(row=1, column=0)
exercise_entry = tk.Entry(root)
exercise_entry.grid(row=1, column=1)

# Weight
tk.Label(root, text="Weight").grid(row=2, column=0)
weight_entry = tk.Entry(root)
weight_entry.grid(row=2, column=1)

# Sets
tk.Label(root, text="Sets").grid(row=3, column=0)
sets_entry = tk.Entry(root)
sets_entry.grid(row=3, column=1)

# Reps
tk.Label(root, text="Reps").grid(row=4, column=0)
reps_entry = tk.Entry(root)
reps_entry.grid(row=4, column=1)

# Distance (for walking/cycling)
tk.Label(root, text="Distance (km)").grid(row=5, column=0)
distance_entry = tk.Entry(root)
distance_entry.grid(row=5, column=1)

# Duration (for walking/cycling)
tk.Label(root, text="Duration (min)").grid(row=6, column=0)
duration_entry = tk.Entry(root)
duration_entry.grid(row=6, column=1)

# Save Button
save_button = tk.Button(root, text="Save", command=lambda: save_data(
    date_entry.get(), exercise_entry.get(), weight_entry.get(), 
    sets_entry.get(), reps_entry.get(), distance_entry.get(), duration_entry.get()))
save_button.grid(row=7, column=0, columnspan=2)

# Generate Graph Button
graph_button = tk.Button(root, text="Generate Graph", command=generate_graph)
graph_button.grid(row=8, column=0, columnspan=2)

# Runkeeper API Token Entry
tk.Label(root, text="Runkeeper API Token").grid(row=9, column=0)
token_entry = tk.Entry(root)
token_entry.grid(row=9, column=1)

# Fetch Data Button
fetch_button = tk.Button(root, text="Fetch Runkeeper Data", command=lambda: fetch_runkeeper_data(token_entry.get()))
fetch_button.grid(row=10, column=0, columnspan=2)

root.mainloop()
