import pandas as pd
import matplotlib.pyplot as plt

# 解决中文字体问题，更换为系统中存在的字体（如SimHei）
plt.rcParams['font.sans-serif'] = ['SimHei']  # 替换为系统支持的中文字体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 读取文件（请根据实际情况修改文件路径）
excel_file = pd.ExcelFile('E:\人力资源课题/318263147_按文本_人力资源服务产业发展之政府支持需求调研问卷_22_22.xlsx')

# 获取指定工作表中的数据
df = excel_file.parse('Sheet1')

# 提取重视度数据列和满意度数据列
importance_cols = df.columns[10:28]
satisfaction_cols = df.columns[28:46]

# 将重视度和满意度的答案进行数值化
importance_mapping = {'非常不重要': 1, '比较不重要': 2, '一般': 3, '比较重要': 4, '非常重要': 5}
satisfaction_mapping = {'非常不好': 1, '比较不好': 2, '一般': 3, '比较好': 4, '非常好': 5}

# 应用数值化映射（明确使用loc避免潜在警告，并处理downcasting警告）
for col in importance_cols:
    df.loc[:, col] = df[col].replace(importance_mapping)
for col in satisfaction_cols:
    df.loc[:, col] = df[col].replace(satisfaction_mapping)

# 处理replace的downcasting警告
df = df.infer_objects()

# 计算各项措施的重视度均值和满意度均值（使用loc明确列标签）
importance_means = df[importance_cols].mean()
satisfaction_means = df[satisfaction_cols].mean()

# 计算整体的重视度均值和满意度均值，作为四象限划分的界限
overall_importance_mean = importance_means.mean()
overall_satisfaction_mean = satisfaction_means.mean()

# 设置图片清晰度
plt.rcParams['figure.dpi'] = 300

# 绘制IPA四象限矩阵图
plt.figure(figsize=(12, 10))
plt.scatter(importance_means, satisfaction_means, s=50, alpha=0.7, color='skyblue')

# 添加措施名称标注，调整位置避免重叠
for i, txt in enumerate(importance_cols):
    y_offset = 30 if satisfaction_means[i] > overall_satisfaction_mean else -20
    plt.annotate(txt, (importance_means[i], satisfaction_means[i]),
                 textcoords="offset points", xytext=(0, y_offset),
                 ha='center', fontsize=8, bbox=dict(facecolor='white', alpha=0.6))

# 绘制四象限划分线
plt.axvline(x=overall_importance_mean, color='r', linestyle='--', alpha=0.8)
plt.axhline(y=overall_satisfaction_mean, color='r', linestyle='--', alpha=0.8)

# 添加轴标签和图标题
plt.xlabel('重视度均值', fontsize=12, fontweight='bold')
plt.ylabel('满意度均值', fontsize=12, fontweight='bold')
plt.title('人力资源服务产业政府支持措施IPA四象限分析图', fontsize=14, fontweight='bold')

# 标注四个标准象限区域
plt.text(overall_importance_mean * 0.5, overall_satisfaction_mean * 1.5, '优势区\n（高重视度-高满意度）',
         horizontalalignment='center', verticalalignment='center',
         fontsize=12, fontweight='bold', color='green', bbox=dict(facecolor='white', alpha=0.8))
plt.text(overall_importance_mean * 1.5, overall_satisfaction_mean * 1.5, '维持区\n（低重视度-高满意度）',
         horizontalalignment='center', verticalalignment='center',
         fontsize=12, fontweight='bold', color='blue', bbox=dict(facecolor='white', alpha=0.8))
plt.text(overall_importance_mean * 0.5, overall_satisfaction_mean * 0.5, '机会区\n（低重视度-低满意度）',
         horizontalalignment='center', verticalalignment='center',
         fontsize=12, fontweight='bold', color='orange', bbox=dict(facecolor='white', alpha=0.8))
plt.text(overall_importance_mean * 1.5, overall_satisfaction_mean * 0.5, '改进区\n（高重视度-低满意度）',
         horizontalalignment='center', verticalalignment='center',
         fontsize=12, fontweight='bold', color='red', bbox=dict(facecolor='white', alpha=0.8))

# 调整坐标轴范围
x_min, x_max = importance_means.min() - 0.2, importance_means.max() + 0.2
y_min, y_max = satisfaction_means.min() - 0.2, satisfaction_means.max() + 0.2
plt.xlim(x_min, x_max)
plt.ylim(y_min, y_max)

# 添加网格线
plt.grid(True, linestyle='--', alpha=0.4)

# 保存图片（请根据需要修改保存路径）
plt.savefig('hr_support_ipa_analysis.png', bbox_inches='tight')

plt.show()

# 生成18项措施的重视度和满意度均值表格
mean_data = {
    '政府支持措施': importance_cols,
    '重视度均值': importance_means.values,
    '满意度均值': satisfaction_means.values
}

mean_df = pd.DataFrame(mean_data)

# 输出表格到控制台
print("\n18项政府支持措施的重视度和满意度均值表格：")
print(mean_df.to_string(index=False))

# 也可以保存为CSV文件
mean_df.to_csv('hr_support_means.csv', index=False, encoding='utf-8-sig')