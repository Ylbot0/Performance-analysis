import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk


def process_excel_files(a1_path, a2_path):
    """
    读取两个Excel文件，按姓名合并并计算成绩变化
    :param a1_path: A1.xlsx文件路径
    :param a2_path: A2.xlsx文件路径
    :return: 处理后的DataFrame，包含姓名及各科成绩变化列
    """
    df1 = pd.read_excel(a1_path)
    df2 = pd.read_excel(a2_path)
    result_df = pd.merge(df2, df1, on='姓名', how='left', suffixes=('_A2', '_A1'))
    result_cols = ['姓名']
    for col in df2.columns[1:]:
        if col in df1.columns:
            change_col_name = col + '_变化'
            result_df[change_col_name] = result_df[col + '_A2'] - result_df[col + '_A1']
            result_cols.append(change_col_name)
    return result_df[result_cols]


def generate_award_list(result_df, award_mode, custom_score=5, selected_subjects=None):
    """
    根据不同模式和条件生成获奖名单数据框
    :param result_df: 处理后的成绩数据框
    :param award_mode: 奖状生成模式（"全部科目"、"特定科目"、"自定义分数"）
    :param custom_score: 自定义分数（用于特定科目和自定义分数模式）
    :param selected_subjects: 特定科目列表（特定科目模式下勾选的科目）
    :return: 获奖名单数据框
    """
    award_df = pd.DataFrame(columns=['姓名', '奖项'])
    subjects = [col[:-3] for col in result_df.columns if col.endswith('_变化')]
    for index, row in result_df.iterrows():
        name = row['姓名']
        for subject in subjects:
            change_col_name = subject + '_变化'
            change = row[change_col_name]
            if award_mode == "全部科目":
                if all(row[col + '_变化'] > 0 for col in subjects):
                    new_row = pd.DataFrame({'姓名': [name], '奖项': '全能进步之星'})
                    award_df = pd.concat([award_df, new_row], ignore_index=True)
            elif award_mode == "特定科目":
                if selected_subjects and subject in selected_subjects and change >= custom_score:
                    new_row = pd.DataFrame({'姓名': [name], '奖项': subject + '进步之星'})
                    award_df = pd.concat([award_df, new_row], ignore_index=True)
            elif award_mode == "自定义分数":
                if change > custom_score:
                    new_row = pd.DataFrame({'姓名': [name], '奖项': subject + '进步之星'})
                    award_df = pd.concat([award_df, new_row], ignore_index=True)
    return award_df


def select_files():
    root = tk.Tk()
    root.title("成绩处理工具")

    # A1.xlsx文件路径相关组件
    a1_path_label = ttk.Label(root, text="A1.xlsx文件路径:")
    a1_path_label.pack()
    a1_path_entry = ttk.Entry(root, width=50)
    a1_path_entry.pack()
    a1_path_button = ttk.Button(root, text="选择文件",
                                command=lambda: get_file_path(a1_path_entry, "选择A1.xlsx文件", "*.xlsx"))
    a1_path_button.pack()

    # A2.xlsx文件路径相关组件
    a2_path_label = ttk.Label(root, text="A2.xlsx文件路径:")
    a2_path_label.pack()
    a2_path_entry = ttk.Entry(root, width=50)
    a2_path_entry.pack()
    a2_path_button = ttk.Button(root, text="选择文件",
                                command=lambda: get_file_path(a2_path_entry, "选择A2.xlsx文件", "*.xlsx"))
    a2_path_button.pack()

    # 保存文件目录路径相关组件
    save_dir_label = ttk.Label(root, text="保存文件目录:")
    save_dir_label.pack()
    save_dir_entry = ttk.Entry(root, width=50)
    save_dir_entry.pack()
    save_dir_button = ttk.Button(root, text="选择目录",
                                 command=lambda: get_dir_path(save_dir_entry))
    save_dir_button.pack()

    # 奖状生成模式选择相关组件
    award_mode_label = ttk.Label(root, text="奖状生成模式:")
    award_mode_label.pack()
    award_mode_var = tk.StringVar()
    award_mode_var.set("全部科目")
    all_subject_radio = ttk.Radiobutton(root, text="全部科目", variable=award_mode_var, value="全部科目")
    all_subject_radio.pack()
    specific_subject_radio = ttk.Radiobutton(root, text="特定科目", variable=award_mode_var, value="特定科目")
    specific_subject_radio.pack()
    custom_score_radio = ttk.Radiobutton(root, text="自定义分数", variable=award_mode_var, value="自定义分数")
    custom_score_radio.pack()

    # 特定科目相关组件
    subject_frame = tk.Frame(root)
    subject_checkbuttons = {}

    # 自定义分数相关组件
    custom_score_label = ttk.Label(root, text="自定义进步分数:")
    custom_score_label.pack_forget()
    custom_score_entry = ttk.Entry(root, width=10)
    custom_score_entry.pack_forget()

    def show_hide_custom_widgets():
        mode = award_mode_var.get()
        if mode == "特定科目":
            subject_frame.pack()
            for subject in ["语文", "数学", "英语", "思品", "物理", "化学", "历史"]:
                var = tk.IntVar()
                checkbutton = ttk.Checkbutton(subject_frame, text=subject, variable=var)
                checkbutton.pack(side=tk.LEFT)
                subject_checkbuttons[subject] = var
        elif mode == "自定义分数":
            custom_score_label.pack()
            custom_score_entry.pack()
        else:
            subject_frame.pack_forget()
            custom_score_label.pack_forget()
            custom_score_entry.pack_forget()

    award_mode_var.trace("w", lambda *args: show_hide_custom_widgets())

    # 定义custom_score为全局变量，并设置一个初始值（可根据需求调整）
    custom_score = 0

    process_button = ttk.Button(root, text="处理文件",
                                command=lambda: process_and_save(a1_path_entry.get(), a2_path_entry.get(),
                                                                 save_dir_entry.get(),
                                                                 award_mode_var.get(),
                                                                 custom_score,
                                                                 [subject for subject, var in subject_checkbuttons.items() if var.get()])
                                )
    process_button.pack()

    # 添加关于按钮相关组件及功能
    about_button = ttk.Button(root, text="关于", command=show_about_popup)
    about_button.pack()

    root.mainloop()


def get_file_path(entry_widget, title, filetypes):
    file_path = filedialog.askopenfilename(title=title, filetypes=[(filetypes.split("*")[1], filetypes)])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)


def get_dir_path(entry_widget):
    dir_path = filedialog.askdirectory(title="选择保存文件目录")
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, dir_path)


def process_and_save(a1_path, a2_path, save_dir, award_mode, custom_score, selected_subjects=None):
    if a1_path and a2_path:
        result_df = process_excel_files(a1_path, a2_path)
        if save_dir:
            change_path = save_dir + "/变化.xlsx"
            result_df.to_excel(change_path, index=False)

            award_df = generate_award_list(result_df, award_mode, custom_score, selected_subjects)

            award_path = save_dir + "/获奖名单.xlsx"
            award_df.to_excel(award_path, index=False)

            # 处理完成后提示并关闭程序
            root = tk.Tk()
            root.title("提示")
            label = ttk.Label(root, text="文件处理完成！")
            label.pack()
            button = ttk.Button(root, text="确定", command=root.destroy)
            button.pack()
            root.mainloop()
    else:
        root = tk.Tk()
        root.title("错误")
        label = ttk.Label(root, text="请填写完整文件路径和保存目录信息！")
        label.pack()
        button = ttk.Button(root, text="确定", command=root.destroy)
        button.pack()
        root.mainloop()


def show_about_popup():
    about_root = tk.Tk()
    about_root.title("关于")

    about_text = ttk.Label(about_root, text="这是一个成绩处理工具，用于处理Excel文件中的成绩数据并生成相应的获奖名单等。开发者：小路 荒稽")
    about_text.pack()

    feedback_button = ttk.Button(about_root, text="反馈", command=open_feedback_page)
    feedback_button.pack()

    about_root.mainloop()


def open_feedback_page():
    import webbrowser
    webbrowser.open("https://www.wjx.cn/vm/mOHJyYO.aspx# ")  # 这里替换为实际的反馈页面网址


if __name__ == "__main__":
    select_files()