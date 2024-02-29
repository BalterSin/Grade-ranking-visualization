import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = 'C:\VSProject\Code\量7.xlsx'  # 请将此处替换为您的Excel文件路径
xls = pd.ExcelFile(file_path)

# 设置Matplotlib以支持中文显示
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 选择一个支持中文的字体
plt.rcParams['axes.unicode_minus'] = False  # 正确显示负号

# Read and process each sheet
rankings_data = {}
sheet_names = xls.sheet_names
for name in sheet_names:
    df = pd.read_excel(xls, sheet_name=name, usecols=["姓名", "排名"])
    rankings_data[name] = df.set_index("姓名")["排名"]

# Combine rankings from all exams into a single DataFrame
combined_rankings = pd.DataFrame(rankings_data)

# Fill NA values with 40 (assuming missing students have a ranking of 40)
combined_rankings_filled = combined_rankings.fillna(41).astype(int)


# Generate a separate line plot for each student and save the plot
for student in combined_rankings_filled.index:
    plt.figure(figsize=(10, 6))
    student_rankings = combined_rankings_filled.loc[student]
    plt.plot(student_rankings.index, student_rankings, marker='o', linestyle='-')
    # Adding text annotation for rankings
    for exam, ranking in student_rankings.items():
        annotation = str(ranking) if ranking != 41 else '缺考'
        plt.text(exam, ranking, annotation, ha='center')
    



    plt.xticks(rotation=45)
    plt.xlabel('考试')
    plt.ylabel('排名')
    plt.title(f'{student}的成绩发展趋势')
    plt.ylim(0, 41)  # Fixing y-axis range from 0 to 40
    plt.gca().invert_yaxis()  # Invert y-axis to have the top rank at the top
    
    # Adjust margins and save the figure
    plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)
    plt.savefig(f'{student}的成绩发展趋势.png')
    plt.close()
