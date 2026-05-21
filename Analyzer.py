import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

sns.set_theme(style="whitegrid")
plt.rcParams['font.family'] = 'DejaVu Sans'

def analyze_performance(file_path):
    try:
        df = pd.read_excel(file_path)
        if "ФИО ученика" not in df.columns:
            raise ValueError("В файле должна быть колонка 'ФИО ученика'")

        grade_columns = [col for col in df.columns if col !="ФИО ученика"]

        # Средний балл каждого ученика 
        df["Средний балл"] = df[grade_columns].mean(axis =1).round(2)

        # Определение динамики (последняя оценка минус среднее по прошлым темам)
        if len(grade_columns)>1:
            prev_topics_avg = df[grade_columns[:-1]].mean(axis=1)
            df["Динамика"] = (df[grade_columns[-1]]- prev_topics_avg).round(2)
        else:
            df["Динамика"] = 0

        risk_zone = df[(df["Средний балл"] < 3.6) | (df["Динамика"]< -0.5)]

        topic_means = df[grade_columns].mean().round(2).reset_index()
        topic_means.columns = ["Тема", "Средний балл класса"]

        is_hard = topic_means["Средний балл класса"] < 3.8
        topic_means["Сложная тема"] = is_hard.map({True: "Да", False: "Нет"})

        output_path = os.path.join(os.path.dirname(file_path), "Анализ Успеваемости Отчет.xlsx")
        with pd.ExcelWriter(output_path, engine = "openpyxl") as writer:
            df.to_excel(writer, sheet_name = "Общая успеваемость", index = False)
            risk_zone.to_excel(writer, sheet_name = "Зона риска", index = False)
            topic_means.to_excel(writer, sheet_name = "Анализ тем", index = False)

        create_plots(df, topic_means, os.path.dirname(file_path))

        return output_path

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при анализе:\n{str(e)}")
        return None 

def create_plots(df, topic_means, output_dir):
    fig, axes = plt.subplots(2, 1, figsize=(10, 12))  

    sns.barplot(
        x="Средний балл",
        y ="ФИО ученика",
        data = df.sort_values("Средний балл", ascending = False),
        ax = axes[0],
        palette ="Blues_r"
    )
    axes[0].set_title("Рейтинг учеников по среднему баллу", fontsize =14, fontweight ="bold")
    axes[0].set_xlabel("Средний балл")
    axes[0].set_ylabel("Ученик")
    axes[0].axvline(3.6, color ="red", linestyle ="--", label ="Граница зоны риска(3.6 балла по теме у ученика )")
    axes[0].legend()

    sns.lineplot(
        x= "Тема",
        y = "Средний балл класса",
        data = topic_means,
        ax = axes[1],
        marker = "o",
        linewidth =2.5,
        color = "teal"
    )
    axes[1].set_title("Анализ усвоения учебных тем", fontsize =14, fontweight ="bold")
    axes[1].set_xlabel("Учебная тема")
    axes[1].set_ylabel("средний балл класса")
    axes[1].set_ylim(1, 5.5)
    axes[1].axhline(3.8, color ="orange", linestyle ="--", label ="Целевой ориентир(3.8 балла )")
    axes[1].legend()

    axes[1].set_xticklabels(topic_means["Тема"], rotation = 15, ha = "right")

    plt.tight_layout()

    plot_path = os.path.join(output_dir, "Аналитика успеваемости.png")
    plt.savefig(plot_path, dpi = 300)
    plt.close()


